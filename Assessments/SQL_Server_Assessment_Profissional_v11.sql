/*
===============================================================================
 SQL SERVER ASSESSMENT / DISCOVERY - V11  (Solo Network)
 Baseline: SQL Server 2017+
 Compatibilidade: SQL Server 2017, 2019, 2022 e 2025
 Objetivo: levantamento técnico não invasivo de uma nova instância.

 - Somente leitura/coleta.
 - Não cria objetos permanentes (apenas tabelas temporárias #).
 - Não altera configurações, bancos, AGs, endpoints ou jobs.
 - SQL Server 2017 (14.x) é a versão mínima suportada.
 - Algumas informações dependem das permissões da conta de execução
   (recomendado: sysadmin ou VIEW SERVER STATE + acesso de leitura à msdb).

 NOVIDADES DA V11
 - Log shipping: configuração das databases primárias e secundárias (pastas,
   retenção, compressão, modo de restore, atraso, servidor monitor), jobs
   LSBackup/LSCopy/LSRestore/LSAlert, último backup/cópia/restore com atraso
   x limite do monitor, erros dos últimos 7 dias e novas verificações no
   checklist (categoria 23). Exibido na guia Alta Disponibilidade.
 - Relatório HTML organizado em guias (Visão Geral, Atenção, Databases,
   Performance, Backups, Segurança, Jobs, Armazenamento, Conexões, Alta
   Disponibilidade e Checklist), gerado diretamente pelo script.

 NOVIDADES DA V10
 - Segurança: permissões explícitas de servidor, membros de server roles e
   política de senha por login SQL (inclui senha em branco / igual ao login).
 - Objetos por database, fragmentação de índices (modo LIMITED, parâmetro
   @ColetarFragmentacao) e missing indexes (top 25).
 - Top 20 queries por duração média e por CPU (plan cache).
 - Sessões de usuário no momento da coleta (bloqueios, transações abertas).
 - Histórico de backups Full/Diff e resumo diário de log (7 dias), com taxa de
   compressão, destino e alerta de backup no mesmo volume dos dados.
 - Agendamentos dos jobs com próxima execução.
 - Novos itens no checklist (categorias 18 a 22) e detalhamentos 19 a 24.

 NOVIDADES DA V9
 - Relatório HTML com a identidade visual Solo Network (último resultset).
 - Novas verificações: ciclo de vida da versão, Instant File Initialization,
   Lock Pages in Memory, max server memory/MAXDOP/CTFP em valor padrão,
   xp_cmdshell, CHECKDB > 7 dias, jobs com falha, alertas de erros graves,
   operadores, logins SQL sem política de senha, uso de log e VLFs.
 - Novo detalhamento: Top 20 wait stats e uso de log por database.
 - Falhas de coleta (permissão, DMV indisponível) não interrompem o script:
   são registradas e exibidas no final do relatório.

 COMO GERAR O HTML
 a) SSMS: execute o script; no ÚLTIMO resultset clique no link da coluna
    "Relatorio_HTML", depois Arquivo > Salvar Como... > nome.html
    (se o conteúdo vier cortado: Ferramentas > Opções > Resultados da Consulta
     > SQL Server > Resultados para Grade > Dados XML = Ilimitado).
 b) PowerShell (recomendado, grava o .html direto):
    .\Gerar-RelatorioAssessment.ps1 -Instancia "SERVIDOR\INSTANCIA" -Cliente "Nome do Cliente"
===============================================================================
*/
USE [master];
GO
SET NOCOUNT ON;
SET XACT_ABORT OFF;   /* uma falha pontual (ex.: permissão) não aborta a coleta inteira */

/* ============================================================================
   PARÂMETROS
============================================================================ */
DECLARE @Modo tinyint = 1;
/*
   MODO DE EXECUÇÃO
   1 = Dashboard + detalhes
   2 = Somente Dashboard
   3 = Somente detalhes
*/
DECLARE @GerarHTML bit = 1;            /* 1 = gera o relatório HTML no último resultset */
DECLARE @SomenteHTML bit = 0;          /* 1 = não retorna os demais resultsets (usado pelo PowerShell) */
DECLARE @Cliente nvarchar(200) = N'';  /* nome do cliente exibido na capa do relatório */
DECLARE @HtmlMaxLinhas int = 500;      /* máximo de linhas por tabela de detalhe no HTML */
DECLARE @ColetarFragmentacao bit = 1;  /* V10: 0 = não coleta fragmentação (instâncias muito grandes / horário de pico) */
DECLARE @FragMinPaginas int = 1000;    /* V10: ignora índices menores que isso (~8 MB) */

DECLARE @ProductMajorVersion int =
    TRY_CONVERT(int, SERVERPROPERTY('ProductMajorVersion'));

DECLARE @ProductVersion nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion'));

DECLARE @ProductLevel nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('ProductLevel'));

DECLARE @Edition nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('Edition'));

DECLARE @IsHadrEnabled bit =
    CONVERT(bit, SERVERPROPERTY('IsHadrEnabled'));

DECLARE @IsClustered bit =
    CONVERT(bit, SERVERPROPERTY('IsClustered'));

DECLARE @ServerName nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('ServerName'));

DECLARE @AssessmentDate datetime =
    GETDATE();

DECLARE @ScriptVersao nvarchar(20) = N'v11';

/* Fim do suporte estendido Microsoft por versão principal (NULL = não mapeado) */
DECLARE @FimSuporte date =
    CASE @ProductMajorVersion
        WHEN 14 THEN '20271012'   /* SQL Server 2017 */
        WHEN 15 THEN '20300108'   /* SQL Server 2019 */
        WHEN 16 THEN '20330111'   /* SQL Server 2022 */
    END;

/*
 Variáveis centrais do Assessment:
 @IsHadrEnabled = Always On/HADR habilitado na instância
 @IsClustered   = FCI (Failover Cluster Instance)
 @ServerName    = nome da instância
 @AssessmentDate= data/hora da coleta
*/
IF @Modo NOT IN (1,2,3)
BEGIN
    RAISERROR('Valor inválido para @Modo. Use 1, 2 ou 3.', 16, 1);
    RETURN;
END;

PRINT '==============================================================';
PRINT ' SQL SERVER ASSESSMENT / DISCOVERY - V11';
PRINT ' SQL Server: ' + ISNULL(@ProductVersion,'N/D');
PRINT ' Product Level: ' + ISNULL(@ProductLevel,'N/D');
PRINT ' Edition: ' + ISNULL(@Edition,'N/D');
PRINT ' Baseline: SQL Server 2017+';
PRINT '==============================================================';

IF @ProductMajorVersion IS NULL OR @ProductMajorVersion < 14
BEGIN
    RAISERROR('Esta versão do Assessment requer SQL Server 2017 (14.x) ou superior.', 16, 1);
    RETURN;
END;

/* ============================================================================
   LIMPEZA DE TEMPORÁRIAS (permite reexecutar na mesma sessão)
============================================================================ */
DROP TABLE IF EXISTS #Dashboard, #AssessmentSummary, #LogSpace, #Secoes, #ColetaErros, #Achados,
    #D01_Instancia, #D02_Hardware, #D02_HostInfo, #D02_Servicos, #D03_ClusterNodes, #D03_ClusterProps,
    #D04_AGs, #D04_Replicas, #D04_Databases, #D04_Cluster, #D05_LinkedServers, #D06_Endpoints,
    #D07_Databases, #D08_Arquivos, #D08_Volumes, #D08_Drives, #D09_TempdbArquivos, #D09_VersionStore,
    #D10_Backups, #D11_Jobs, #D11_Alertas, #D12_Logins, #D13_Config, #D13_ConfigPendente,
    #D14_ProcessMemory, #D14_SysMemory, #D14_PLE, #D14_CPU, #D15_IOLatencia, #D15_IOPorDB,
    #IOWarningResults, #D16_SuspectPages, #D16_TraceFlags, #D17_BufferPool, #D17_VLF, #D18_Waits,
    #D19_Permissoes, #D19_ServerRoles, #D19_PoliticaSenha, #D20_Objetos, #V10_FragBase, #D20_Fragmentacao,
    #D20_FragResumo, #D20_MissingIndexes, #D21_TopDuracao, #D21_TopCPU, #D22_Sessoes, #D23_BackupHist,
    #D23_BackupLog, #D24_JobAgenda, #V11_LSJob, #D25_LSPrimario, #D25_LSSecundario, #D25_LSMonitor,
    #D25_LSJobs, #D25_LSErros, #V11_HtmlDet;

/* Registro das seções de detalhe (usado para exibir no SSMS e montar o HTML) */
CREATE TABLE #Secoes
(
    Ordem int IDENTITY(1,1) PRIMARY KEY,
    Grupo nvarchar(10) NOT NULL,
    Titulo nvarchar(200) NOT NULL,
    Tabela sysname NULL,             /* tabela temporária com os dados */
    OrdenarPor nvarchar(400) NULL,   /* ORDER BY da exibição */
    Nota nvarchar(2000) NULL,        /* seção apenas informativa (sem tabela) */
    Descricao nvarchar(1000) NULL,
    ColunasLargas nvarchar(400) NULL /* colunas com quebra de linha no HTML, separadas por vírgula */
);

/* Falhas de coleta (permissão, DMV indisponível etc.) */
CREATE TABLE #ColetaErros
(
    ID int IDENTITY(1,1) PRIMARY KEY,
    Secao nvarchar(200) NOT NULL,
    Mensagem nvarchar(4000) NULL
);

/* ============================================================================
   INDICADORES CENTRAIS (usados no Dashboard e no HTML)
============================================================================ */
DECLARE @Dashboard_Hadr bit = CONVERT(bit, SERVERPROPERTY('IsHadrEnabled'));
DECLARE @Dashboard_Clustered bit = CONVERT(bit, SERVERPROPERTY('IsClustered'));

DECLARE @Dashboard_ServerName nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('ServerName'));
DECLARE @Dashboard_InstanceName nvarchar(128) =
    COALESCE(CONVERT(nvarchar(128), SERVERPROPERTY('InstanceName')), 'MSSQLSERVER');
DECLARE @Dashboard_ProductVersion nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion'));

DECLARE @Dashboard_ProductMajorVersion int =
    CONVERT(int, SERVERPROPERTY('ProductMajorVersion'));

DECLARE @Dashboard_SQLServerVersion nvarchar(128) =
    CASE CONVERT(int, SERVERPROPERTY('ProductMajorVersion'))
        WHEN 17 THEN 'SQL Server 2025'
        WHEN 16 THEN 'SQL Server 2022'
        WHEN 15 THEN 'SQL Server 2019'
        WHEN 14 THEN 'SQL Server 2017'
        WHEN 13 THEN 'SQL Server 2016'
        WHEN 12 THEN 'SQL Server 2014'
        WHEN 11 THEN 'SQL Server 2012'
        WHEN 10 THEN 'SQL Server 2008'
        ELSE 'SQL Server (versão não mapeada)'
    END;

DECLARE @Dashboard_ProductBuild nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('ProductBuild'));
DECLARE @Dashboard_Edition nvarchar(128) =
    CONVERT(nvarchar(128), SERVERPROPERTY('Edition'));

DECLARE @Dashboard_AGs int =
    (SELECT COUNT(*) FROM sys.availability_groups);
DECLARE @Dashboard_Replicas int =
    (SELECT COUNT(*) FROM sys.availability_replicas);
DECLARE @Dashboard_Databases int =
    (SELECT COUNT(*) FROM sys.databases WHERE database_id > 4);
DECLARE @Dashboard_LinkedServers int =
    (SELECT COUNT(*) FROM sys.servers WHERE is_linked = 1 AND is_system = 0);
DECLARE @Dashboard_TempdbFiles int =
    (SELECT COUNT(*) FROM tempdb.sys.database_files WHERE type_desc = 'ROWS');
DECLARE @Dashboard_Jobs int =
    (SELECT COUNT(*) FROM msdb.dbo.sysjobs);
DECLARE @Dashboard_Sysadmin int =
    (SELECT COUNT(*) FROM sys.server_principals
     WHERE IS_SRVROLEMEMBER('sysadmin', name) = 1);

DECLARE @Dashboard_AgUnhealthy int =
    (SELECT COUNT(*)
     FROM sys.dm_hadr_availability_replica_states
     WHERE group_id IS NOT NULL
       AND synchronization_health_desc <> 'HEALTHY');

DECLARE @Dashboard_DatabasesNotOnline int =
    (SELECT COUNT(*) FROM sys.databases
     WHERE database_id > 4 AND state_desc <> 'ONLINE');

DECLARE @Dashboard_AutoShrink int =
    (SELECT COUNT(*) FROM sys.databases
     WHERE database_id > 4 AND is_auto_shrink_on = 1);

DECLARE @Dashboard_AutoClose int =
    (SELECT COUNT(*) FROM sys.databases
     WHERE database_id > 4 AND is_auto_close_on = 1);

/* V9: somente eventos ativos (1=823/824, 2=checksum, 3=torn page).
   4, 5 e 7 são páginas já restauradas/reparadas/desalocadas. */
DECLARE @Dashboard_SuspectPages int =
    (SELECT COUNT(*) FROM msdb.dbo.suspect_pages
     WHERE event_type IN (1,2,3));

DECLARE @Dashboard_LogicalCPU int =
    (SELECT cpu_count FROM sys.dm_os_sys_info);

DECLARE @Dashboard_PhysicalMemoryMB bigint =
    (SELECT physical_memory_kb / 1024 FROM sys.dm_os_sys_info);

DECLARE @Dashboard_VirtualMachine nvarchar(60) =
    (SELECT virtual_machine_type_desc FROM sys.dm_os_sys_info);

DECLARE @Dashboard_HostPlatform nvarchar(128) = N'N/D';
DECLARE @Dashboard_HostDistribution nvarchar(256) = N'N/D';
DECLARE @Dashboard_HostRelease nvarchar(128) = N'N/D';

SELECT
    @Dashboard_HostPlatform = COALESCE(host_platform, N'N/D'),
    @Dashboard_HostDistribution = COALESCE(host_distribution, N'N/D'),
    @Dashboard_HostRelease = COALESCE(host_release, N'N/D')
FROM sys.dm_os_host_info;

DECLARE @Dashboard_StatusAtencao int =
      CASE WHEN @Dashboard_AgUnhealthy > 0 THEN 1 ELSE 0 END
    + CASE WHEN @Dashboard_AutoShrink > 0 THEN 1 ELSE 0 END
    + CASE WHEN @Dashboard_AutoClose > 0 THEN 1 ELSE 0 END
    + CASE WHEN @Dashboard_TempdbFiles = 1 THEN 1 ELSE 0 END;

DECLARE @Dashboard_StatusCritico int =
      CASE WHEN @Dashboard_DatabasesNotOnline > 0 THEN 1 ELSE 0 END
    + CASE WHEN @Dashboard_SuspectPages > 0 THEN 1 ELSE 0 END;

/* Indicadores adicionais para o relatório HTML */
DECLARE @K_Sockets int, @K_CoresPorSocket int, @K_Numa int, @K_StartTime datetime;
SELECT @K_Sockets = socket_count, @K_CoresPorSocket = cores_per_socket,
       @K_Numa = numa_node_count, @K_StartTime = sqlserver_start_time
FROM sys.dm_os_sys_info;

DECLARE @K_MaxServerMemoryMB bigint =
    (SELECT CONVERT(bigint, value_in_use) FROM sys.configurations WHERE name = 'max server memory (MB)');
DECLARE @K_DataGB decimal(18,2) =
    (SELECT SUM(CONVERT(bigint, size)) / 128.0 / 1024.0 FROM sys.master_files WHERE type = 0);
DECLARE @K_LogGB decimal(18,2) =
    (SELECT SUM(CONVERT(bigint, size)) / 128.0 / 1024.0 FROM sys.master_files WHERE type = 1);
DECLARE @K_JobsFalha int = 0;
BEGIN TRY
    SELECT @K_JobsFalha = COUNT(*)
    FROM msdb.dbo.sysjobs j
    CROSS APPLY
    (
        SELECT TOP (1) h.run_status
        FROM msdb.dbo.sysjobhistory h
        WHERE h.job_id = j.job_id AND h.step_id = 0
        ORDER BY h.instance_id DESC
    ) lh
    WHERE j.enabled = 1 AND lh.run_status = 0;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'Indicadores - Jobs com falha', ERROR_MESSAGE());
END CATCH;


/* V11: Log shipping (indicador da capa e da guia Alta Disponibilidade) */
DECLARE @K_LSPrim int = 0, @K_LSSec int = 0, @K_LSMon int = 0;
BEGIN TRY
    SELECT @K_LSPrim = (SELECT COUNT(*) FROM msdb.dbo.log_shipping_primary_databases),
           @K_LSSec  = (SELECT COUNT(*) FROM msdb.dbo.log_shipping_secondary_databases),
           @K_LSMon  = (SELECT COUNT(*) FROM msdb.dbo.log_shipping_monitor_primary)
                     + (SELECT COUNT(*) FROM msdb.dbo.log_shipping_monitor_secondary);
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'Indicadores - Log shipping', ERROR_MESSAGE());
END CATCH;

IF @Modo IN (1,2)
BEGIN
/* ============================================================================
   DASHBOARD EXECUTIVO - RESULTSET 01
   IDENTIFICAÇÃO | SISTEMA OPERACIONAL | TOPOLOGIA | RECURSOS | ASSESSMENT
============================================================================= */
SELECT *
INTO #Dashboard
FROM
(
SELECT
    '01' AS ID, 'IDENTIFICAÇÃO' AS Categoria, 'Servidor' AS Indicador,
    @Dashboard_ServerName AS Valor, 'INFO' AS Status,
    'Nome do servidor físico/virtual.' AS Observacao
UNION ALL SELECT
    '02', 'IDENTIFICAÇÃO', 'Instância', @Dashboard_InstanceName, 'INFO',
    'Nome da instância.'
UNION ALL SELECT
    '03', 'IDENTIFICAÇÃO', 'SQL Server', @Dashboard_SQLServerVersion, 'INFO',
    'Versão principal do produto. O build completo é apresentado na linha seguinte.'
UNION ALL SELECT
    '04', 'IDENTIFICAÇÃO', 'Build',
    @Dashboard_ProductVersion + ' (Build ' + @Dashboard_ProductBuild + ')',
    'INFO',
    'Versão completa e build do SQL Server.'
UNION ALL SELECT
    '05', 'IDENTIFICAÇÃO', 'Edition', @Dashboard_Edition, 'INFO',
    'Edição instalada.'
UNION ALL SELECT
    '06', 'SISTEMA OPERACIONAL', 'Plataforma', @Dashboard_HostPlatform, 'INFO',
    'Plataforma reportada pelo sistema operacional.'
UNION ALL SELECT
    '07', 'SISTEMA OPERACIONAL', 'Distribuição', @Dashboard_HostDistribution, 'INFO',
    'Nome/distribuição do sistema operacional.'
UNION ALL SELECT
    '08', 'SISTEMA OPERACIONAL', 'Release', @Dashboard_HostRelease, 'INFO',
    'Release do sistema operacional.'
UNION ALL SELECT
    '09', 'SISTEMA OPERACIONAL', 'Virtualização',
    COALESCE(@Dashboard_VirtualMachine, 'NONE'), 'INFO',
    'Tipo de virtualização reportado pelo SQL Server.'
UNION ALL SELECT
    '10', 'TOPOLOGIA', 'FCI / Cluster',
    CASE WHEN @Dashboard_Clustered = 1 THEN 'SIM' ELSE 'NÃO' END,
    'INFO',
    CASE WHEN @Dashboard_Clustered = 1
         THEN 'Instância configurada como Failover Cluster Instance.'
         ELSE 'FCI não configurado para esta instância.' END
UNION ALL SELECT
    '11', 'TOPOLOGIA', 'Always On / HADR',
    CASE WHEN @Dashboard_Hadr = 1 THEN 'HABILITADO' ELSE 'NÃO HABILITADO' END,
    'INFO', 'Configuração HADR da instância.'
UNION ALL SELECT
    '12', 'TOPOLOGIA', 'Availability Groups',
    CONVERT(nvarchar(30), @Dashboard_AGs), 'INFO',
    'Quantidade de Availability Groups.'
UNION ALL SELECT
    '13', 'TOPOLOGIA', 'Réplicas',
    CONVERT(nvarchar(30), @Dashboard_Replicas), 'INFO',
    'Quantidade de réplicas configuradas.'
UNION ALL SELECT
    '14', 'TOPOLOGIA', 'Linked Servers',
    CONVERT(nvarchar(30), @Dashboard_LinkedServers), 'INFO',
    'Quantidade de Linked Servers não sistêmicos.'
UNION ALL SELECT
    '15', 'RECURSOS', 'CPU',
    CONVERT(nvarchar(30), @Dashboard_LogicalCPU), 'INFO',
    'CPUs lógicas visíveis para o SQL Server.'
UNION ALL SELECT
    '16', 'RECURSOS', 'Memória',
    CONVERT(nvarchar(30), @Dashboard_PhysicalMemoryMB / 1024) + ' GB',
    'INFO', 'Memória física reportada pelo SQL Server.'
UNION ALL SELECT
    '17', 'RECURSOS', 'Databases',
    CONVERT(nvarchar(30), @Dashboard_Databases), 'INFO',
    'Quantidade de databases de usuário.'
UNION ALL SELECT
    '18', 'RECURSOS', 'TempDB - arquivos de dados',
    CONVERT(nvarchar(30), @Dashboard_TempdbFiles),
    CASE WHEN @Dashboard_TempdbFiles = 1 THEN 'ATENÇÃO' ELSE 'INFO' END,
    CASE WHEN @Dashboard_TempdbFiles = 1
         THEN 'Apenas um arquivo de dados no TempDB; investigar dimensionamento.'
         ELSE 'Quantidade de arquivos de dados do TempDB.' END
UNION ALL SELECT
    '19', 'RECURSOS', 'SQL Agent Jobs',
    CONVERT(nvarchar(30), @Dashboard_Jobs), 'INFO',
    'Quantidade total de jobs cadastrados.'
UNION ALL SELECT
    '20', 'RECURSOS', 'Logins com sysadmin',
    CONVERT(nvarchar(30), @Dashboard_Sysadmin), 'INFO',
    'Quantidade de principais associados ao sysadmin.'
UNION ALL SELECT
    '21', 'ASSESSMENT', 'AGs não saudáveis',
    CONVERT(nvarchar(30), @Dashboard_AgUnhealthy),
    CASE WHEN @Dashboard_AgUnhealthy > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
    'Estados HADR que requerem investigação.'
UNION ALL SELECT
    '22', 'ASSESSMENT', 'Databases fora de ONLINE',
    CONVERT(nvarchar(30), @Dashboard_DatabasesNotOnline),
    CASE WHEN @Dashboard_DatabasesNotOnline > 0 THEN 'CRÍTICO' ELSE 'OK' END,
    'Databases de usuário que não estão ONLINE.'
UNION ALL SELECT
    '23', 'ASSESSMENT', 'Auto Shrink',
    CONVERT(nvarchar(30), @Dashboard_AutoShrink),
    CASE WHEN @Dashboard_AutoShrink > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
    'Databases de usuário com AUTO_SHRINK habilitado.'
UNION ALL SELECT
    '24', 'ASSESSMENT', 'Auto Close',
    CONVERT(nvarchar(30), @Dashboard_AutoClose),
    CASE WHEN @Dashboard_AutoClose > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
    'Databases de usuário com AUTO_CLOSE habilitado.'
UNION ALL SELECT
    '25', 'ASSESSMENT', 'Suspect Pages',
    CONVERT(nvarchar(30), @Dashboard_SuspectPages),
    CASE WHEN @Dashboard_SuspectPages > 0 THEN 'CRÍTICO' ELSE 'OK' END,
    'Registros ativos (event_type 1, 2 e 3) em msdb.dbo.suspect_pages.'
UNION ALL SELECT
    '26', 'ASSESSMENT', 'Total de ATENÇÕES',
    CONVERT(nvarchar(30), @Dashboard_StatusAtencao),
    CASE WHEN @Dashboard_StatusAtencao > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
    'Quantidade de critérios automáticos de atenção.'
UNION ALL SELECT
    '27', 'ASSESSMENT', 'Total de CRÍTICOS',
    CONVERT(nvarchar(30), @Dashboard_StatusCritico),
    CASE WHEN @Dashboard_StatusCritico > 0 THEN 'CRÍTICO' ELSE 'OK' END,
    'Quantidade de critérios automáticos críticos.'
) AS D;

IF @SomenteHTML = 0
    SELECT ID, Categoria, Indicador, Valor, Status, Observacao
    FROM #Dashboard
    ORDER BY ID;

PRINT '';
PRINT '==============================================================';
PRINT ' RESULTADO 01 = DASHBOARD EXECUTIVO';
PRINT ' IDENTIFICAÇÃO | SISTEMA OPERACIONAL | TOPOLOGIA | RECURSOS | ASSESSMENT';
PRINT ' Os próximos resultsets apresentam os detalhes do ambiente.';
PRINT '==============================================================';
PRINT '';

END;

IF @Modo IN (1,3)
BEGIN
/* ============================================================================
   TABELAS TEMPORÁRIAS CENTRAIS
============================================================================ */
CREATE TABLE #AssessmentSummary
(
    ID int IDENTITY(1,1) PRIMARY KEY,
    Categoria nvarchar(80) NOT NULL,
    Indicador nvarchar(150) NOT NULL,
    Valor nvarchar(500) NULL,
    Status nvarchar(20) NOT NULL,
    Observacao nvarchar(1000) NULL
);

CREATE TABLE #LogSpace
(
    DatabaseName sysname,
    LogSizeMB decimal(18,2),
    LogSpaceUsedPct decimal(9,2),
    Status int
);

/* ============================================================================
   01 - INSTÂNCIA / VERSÃO
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '01 - Instância', 'Servidor', ISNULL(CONVERT(varchar(128), SERVERPROPERTY('ComputerNamePhysicalNetBIOS')), @@SERVERNAME), 'INFO', 'Nome do servidor físico/virtual.'
UNION ALL SELECT '01 - Instância', 'Instância', ISNULL(CONVERT(varchar(128), SERVERPROPERTY('InstanceName')), 'MSSQLSERVER'), 'INFO', 'Nome da instância.'
UNION ALL SELECT '01 - Instância', 'SQL Server', CONVERT(varchar(128), SERVERPROPERTY('ProductVersion')), 'INFO', 'Versão/build.'
UNION ALL SELECT '01 - Instância', 'Edition', CONVERT(varchar(128), SERVERPROPERTY('Edition')), 'INFO', 'Edição instalada.'
UNION ALL SELECT '01 - Instância', 'Product Level', CONVERT(varchar(128), SERVERPROPERTY('ProductLevel')), 'INFO', 'RTM/SP/branch.'
UNION ALL SELECT '01 - Instância', 'Update Level', ISNULL(CONVERT(varchar(128), SERVERPROPERTY('ProductUpdateLevel')), 'N/D'), 'INFO', 'CU/GDR quando disponível.'
UNION ALL SELECT '01 - Instância', 'KB/Update Reference', ISNULL(CONVERT(varchar(128), SERVERPROPERTY('ProductUpdateReference')), 'N/D'), 'INFO', 'Referência do update.'
UNION ALL SELECT '01 - Instância', 'Collation', CONVERT(varchar(128), SERVERPROPERTY('Collation')), 'INFO', 'Collation da instância.'
UNION ALL SELECT '01 - Instância', 'Autenticação', CASE WHEN SERVERPROPERTY('IsIntegratedSecurityOnly') = 1 THEN 'Windows Only' ELSE 'Windows + SQL Server' END, 'INFO', 'Modo de autenticação.'
UNION ALL SELECT '01 - Instância', 'Always On habilitado', CASE WHEN @IsHadrEnabled = 1 THEN 'SIM' ELSE 'NÃO' END, 'INFO', 'HADR habilitado na instância.'
UNION ALL SELECT '01 - Instância', 'FCI / Failover Cluster Instance', CASE WHEN @IsClustered = 1 THEN 'SIM' ELSE 'NÃO' END, 'INFO', 'SERVERPROPERTY(IsClustered). Não confundir com WSFC usado por Always On.';

/* V9 - ciclo de vida e uptime */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'01 - Instância', N'Ciclo de vida (fim do suporte estendido)',
       ISNULL(CONVERT(nvarchar(10), @FimSuporte, 103), N'N/D'),
       CASE WHEN @FimSuporte IS NULL THEN N'INFO'
            WHEN @FimSuporte < CONVERT(date, GETDATE()) THEN N'CRÍTICO'
            WHEN @FimSuporte < DATEADD(month, 18, GETDATE()) THEN N'ATENÇÃO'
            ELSE N'OK' END,
       CASE WHEN @FimSuporte IS NULL THEN N'Versão sem data mapeada no script; consultar o Microsoft Lifecycle.'
            WHEN @FimSuporte < CONVERT(date, GETDATE()) THEN N'Versão fora do suporte estendido da Microsoft; sem correções de segurança.'
            WHEN @FimSuporte < DATEADD(month, 18, GETDATE()) THEN N'Menos de 18 meses de suporte estendido; planejar upgrade.'
            ELSE N'Versão dentro do ciclo de suporte da Microsoft.' END
UNION ALL
SELECT N'01 - Instância', N'Uptime (dias)', CONVERT(nvarchar(30), DATEDIFF(day, @K_StartTime, GETDATE())), N'INFO',
       N'Dias desde o último início do serviço. DMVs cumulativas (waits, I/O, índices) refletem este período.';

/* ============================================================================
   02 - SO / HARDWARE / SERVIÇOS
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '02 - Hardware', 'CPU lógica', CONVERT(varchar(30), cpu_count), 'INFO', 'CPU lógica visível ao SQL Server.'
FROM sys.dm_os_sys_info
UNION ALL
SELECT '02 - Hardware', 'Memória física (GB)', CONVERT(varchar(30), CONVERT(decimal(18,2), physical_memory_kb/1024.0/1024.0)), 'INFO', 'Memória física reportada pelo SQL Server.'
FROM sys.dm_os_sys_info
UNION ALL
SELECT '02 - Hardware', 'NUMA nodes', CONVERT(varchar(30), numa_node_count), 'INFO', 'Quantidade de nós NUMA.'
FROM sys.dm_os_sys_info
UNION ALL
SELECT '02 - Hardware', 'Virtualização', ISNULL(virtual_machine_type_desc,'UNKNOWN'), 'INFO', 'Indicador de virtualização do SQL Server; deve ser validado com o SO/hipervisor.'
FROM sys.dm_os_sys_info
UNION ALL
SELECT '02 - Hardware', 'SQL Server Start Time', CONVERT(varchar(30), sqlserver_start_time, 120), 'INFO', 'Último início do serviço SQL Server.'
FROM sys.dm_os_sys_info
UNION ALL
SELECT N'02 - Hardware', N'Lock Pages in Memory', CASE WHEN sql_memory_model_desc = 'CONVENTIONAL' THEN N'NÃO' ELSE N'SIM (' + sql_memory_model_desc + N')' END,
       N'INFO', N'Modelo de memória do SQL Server. Avaliar LPIM em servidores dedicados com max server memory configurado.'
FROM sys.dm_os_sys_info;

BEGIN TRY
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT TOP (1) N'02 - Hardware', N'Instant File Initialization',
           CASE WHEN instant_file_initialization_enabled = 'Y' THEN N'SIM' ELSE N'NÃO' END,
           CASE WHEN instant_file_initialization_enabled = 'Y' THEN N'OK' ELSE N'ATENÇÃO' END,
           N'Sem IFI, crescimentos e restores de arquivos de dados são mais lentos (conta de serviço sem "Perform Volume Maintenance Tasks").'
    FROM sys.dm_server_services
    WHERE servicename LIKE N'SQL Server (%';
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'02 - Instant File Initialization', ERROR_MESSAGE());
END CATCH;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '02 - SO', 'Plataforma', CONCAT(host_platform, ' / ', ISNULL(host_distribution,'')), 'INFO', 'Informação reportada pelo SQL Server.'
FROM sys.dm_os_host_info
UNION ALL
SELECT '02 - SO', 'Release', ISNULL(host_release,'N/D'), 'INFO', 'Release do sistema operacional.'
FROM sys.dm_os_host_info
UNION ALL
SELECT '02 - SO', 'Idioma', CONVERT(varchar(30), os_language_version), 'INFO', 'Idioma do SO.'
FROM sys.dm_os_host_info;

/* ============================================================================
   03 - CLUSTER / FCI
============================================================================ */
IF @IsClustered = 1
BEGIN
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT '03 - Cluster / FCI', 'FCI', 'SIM', 'INFO', 'A instância está instalada como Failover Cluster Instance.';

    BEGIN TRY
        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT '03 - Cluster / FCI', 'Nó do cluster', ISNULL(NodeName,'N/D'), 'INFO', 'Nó identificado pelo DMV de cluster.'
        FROM sys.dm_os_cluster_nodes;
    END TRY
    BEGIN CATCH
        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao) VALUES ('03 - Cluster / FCI','Consulta de nós','Erro ao consultar DMV','ATENÇÃO',ERROR_MESSAGE());
    END CATCH;
END
ELSE
BEGIN
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao) VALUES ('03 - Cluster / FCI','FCI','NÃO','INFO','A instância não está instalada como FCI. Um WSFC pode existir para Always On mesmo assim.');
END;

/* ============================================================================
   04 - ALWAYS ON
============================================================================ */
IF @IsHadrEnabled = 1
BEGIN
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT '04 - Always On', 'Availability Groups', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Quantidade de AGs configurados nesta instância.'
    FROM sys.availability_groups;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT '04 - Always On', 'Réplicas configuradas', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Quantidade de réplicas conhecidas pela instância.'
    FROM sys.availability_replicas;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT '04 - Always On', 'Bancos em AG', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Quantidade de bancos associados a AGs.'
    FROM sys.availability_databases_cluster;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT '04 - Always On', 'Bancos não sincronizados', CONVERT(varchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
           'Contagem baseada no estado reportado pelas réplicas locais/conhecidas.'
    FROM sys.dm_hadr_database_replica_states
    WHERE synchronization_state_desc <> 'SYNCHRONIZED' OR synchronization_state_desc IS NULL;
END
ELSE
BEGIN
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao) VALUES ('04 - Always On','Always On','NÃO HABILITADO','INFO','HADR não está habilitado nesta instância.');
END;

/* ============================================================================
   05 - LINKED SERVERS
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '05 - Linked Servers', 'Quantidade', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Linked Servers configurados, excluindo a própria instância.'
FROM sys.servers
WHERE is_linked = 1;

/* ============================================================================
   06 - ENDPOINTS
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '06 - Endpoints', 'Quantidade', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Endpoints configurados na instância.'
FROM sys.endpoints;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '06 - Endpoints', 'Database Mirroring / HADR endpoint', CONVERT(varchar(30), COUNT(*)),
       'INFO',
       'Endpoint utilizado normalmente em cenários de Always On/Database Mirroring.'
FROM sys.endpoints
WHERE type_desc = 'DATABASE_MIRRORING';

/* ============================================================================
   07 - DATABASES
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '07 - Databases', 'Total de databases', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Inclui bancos de sistema.'
FROM sys.databases;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '07 - Databases', 'Databases de usuário', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Bancos com database_id > 4.'
FROM sys.databases
WHERE database_id > 4;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'07 - Databases', N'Databases fora de ONLINE', CONVERT(nvarchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
       N'Databases de usuário em estado diferente de ONLINE (OFFLINE, RESTORING, SUSPECT, RECOVERY_PENDING...).'
FROM sys.databases WHERE database_id > 4 AND state_desc <> 'ONLINE';

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '07 - Databases', 'Auto Close habilitado', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Auto Close geralmente merece validação em ambientes de produção.'
FROM sys.databases WHERE database_id > 4 AND is_auto_close_on = 1;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '07 - Databases', 'Auto Shrink habilitado', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Auto Shrink pode gerar operações recorrentes de shrink; validar necessidade.'
FROM sys.databases WHERE database_id > 4 AND is_auto_shrink_on = 1;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '07 - Databases', 'Page Verify diferente de CHECKSUM', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Validar bancos que não usam CHECKSUM.'
FROM sys.databases WHERE database_id > 4 AND page_verify_option_desc <> 'CHECKSUM';

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'07 - Databases', N'Compatibility level inferior ao da instância', CONVERT(nvarchar(30), COUNT(*)), N'INFO',
       N'Pode ser intencional (homologação da aplicação); considerar no planejamento de upgrade/otimizações.'
FROM sys.databases WHERE database_id > 4 AND compatibility_level < @ProductMajorVersion * 10;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '07 - Databases', 'Crescimento percentual em arquivos', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Arquivos com growth em percentual merecem revisão.'
FROM sys.master_files
WHERE is_percent_growth = 1;

/* ============================================================================
   08 - ARQUIVOS / VOLUMES / STORAGE
============================================================================ */
BEGIN TRY
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT '08 - Storage', 'Volumes com espaço livre < 15%', CONVERT(varchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
           'Baseado nos volumes que possuem arquivos do SQL Server.'
    FROM
    (
        SELECT DISTINCT vs.volume_mount_point,
               CAST(vs.available_bytes * 100.0 / NULLIF(vs.total_bytes,0) AS decimal(9,2)) AS FreePct
        FROM sys.master_files mf
        CROSS APPLY sys.dm_os_volume_stats(mf.database_id,mf.file_id) vs
    ) V
    WHERE FreePct < 15;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'08 - Volumes (resumo)', ERROR_MESSAGE());
END CATCH;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '08 - Storage', 'Arquivos em C:', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'INFO' END,
       'Validar se armazenamento do SQL Server está concentrado no volume do SO.'
FROM sys.master_files
WHERE LEFT(physical_name,2) = 'C:';

/* ============================================================================
   09 - TEMPDB
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '09 - TempDB', 'Arquivos de dados', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) = 1 THEN 'ATENÇÃO' ELSE 'INFO' END,
       'Um único arquivo não é automaticamente um problema; avaliar contenção e carga.'
FROM tempdb.sys.database_files
WHERE type = 0;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '09 - TempDB', 'Arquivos com tamanhos diferentes', CONVERT(varchar(30), CASE WHEN COUNT(DISTINCT size) > 1 THEN 1 ELSE 0 END),
       CASE WHEN COUNT(DISTINCT size) > 1 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Idealmente os arquivos de dados devem ter tamanho inicial e crescimento coerentes.'
FROM tempdb.sys.database_files
WHERE type = 0;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '09 - TempDB', 'Arquivos com growth percentual', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Validar configuração de crescimento.'
FROM tempdb.sys.database_files
WHERE is_percent_growth = 1;

/* ============================================================================
   10 - BACKUPS / INTEGRIDADE
============================================================================ */
/* DBCC SQLPERF(LOGSPACE) é somente leitura e fornece log size/uso para todos os DBs. */
BEGIN TRY
    INSERT INTO #LogSpace EXEC ('DBCC SQLPERF(LOGSPACE) WITH NO_INFOMSGS');
END TRY
BEGIN CATCH
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao) VALUES ('10 - Backups','DBCC SQLPERF(LOGSPACE)','Não foi possível coletar','ATENÇÃO',ERROR_MESSAGE());
END CATCH;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '10 - Backups', 'DBs em FULL sem backup de log nos últimos 24h', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Avaliação indicativa; validar política de backup e RPO.'
FROM sys.databases d
WHERE d.database_id > 4
  AND d.recovery_model_desc = 'FULL'
  AND NOT EXISTS
  (
      SELECT 1
      FROM msdb.dbo.backupset b
      WHERE b.database_name = d.name
        AND b.type = 'L'
        AND b.backup_finish_date >= DATEADD(hour,-24,GETDATE())
  );

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '10 - Backups', 'DBs sem backup FULL nos últimos 7 dias', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Validar política de backup, retenção e necessidade do banco.'
FROM sys.databases d
WHERE d.database_id > 4
  AND d.state_desc = 'ONLINE'
  AND NOT EXISTS
  (
      SELECT 1
      FROM msdb.dbo.backupset b
      WHERE b.database_name = d.name
        AND b.type = 'D'
        AND b.backup_finish_date >= DATEADD(day,-7,GETDATE())
  );

/* V9 - último CHECKDB com sucesso (1900-01-01 = nunca executado) */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'10 - Backups', N'DBs sem CHECKDB com sucesso nos últimos 7 dias', CONVERT(nvarchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'Baseado em DATABASEPROPERTYEX(LastGoodCheckDbTime). Integridade física/lógica sem verificação recente.'
FROM sys.databases d
CROSS APPLY (SELECT CONVERT(datetime, DATABASEPROPERTYEX(d.name, 'LastGoodCheckDbTime')) AS LastGood) c
WHERE d.database_id <> 2
  AND d.state_desc = 'ONLINE'
  AND c.LastGood IS NOT NULL
  AND c.LastGood < DATEADD(day, -7, GETDATE());

/* ============================================================================
   11 - SQL AGENT / ALERTAS
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '11 - SQL Agent', 'Jobs totais', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Jobs cadastrados no SQL Agent.'
FROM msdb.dbo.sysjobs;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '11 - SQL Agent', 'Jobs desabilitados', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'INFO' ELSE 'OK' END,
       'Revisar se jobs desabilitados são intencionais.'
FROM msdb.dbo.sysjobs WHERE enabled = 0;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'11 - SQL Agent', N'Jobs habilitados com falha na última execução', CONVERT(nvarchar(30), @K_JobsFalha),
       CASE WHEN @K_JobsFalha > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'Ver detalhamento de jobs (coluna LastRunStatusDesc).';

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '11 - SQL Agent', 'Alertas configurados', CONVERT(varchar(30), COUNT(*)), 'INFO', 'Alertas do SQL Agent.'
FROM msdb.dbo.sysalerts;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'11 - SQL Agent', N'Alertas para erros graves (sev. 19-25 / 823-825)', CONVERT(nvarchar(30), COUNT(*)),
       CASE WHEN COUNT(*) = 0 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'Boa prática: alertas habilitados para severidades 19 a 25 e erros 823, 824 e 825 (I/O).'
FROM msdb.dbo.sysalerts
WHERE enabled = 1 AND (severity BETWEEN 19 AND 25 OR message_id IN (823,824,825));

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'11 - SQL Agent', N'Operadores habilitados', CONVERT(nvarchar(30), COUNT(*)),
       CASE WHEN COUNT(*) = 0 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'Sem operadores, falhas de jobs e alertas não geram notificação.'
FROM msdb.dbo.sysoperators
WHERE enabled = 1;

/* ============================================================================
   12 - SEGURANÇA / LOGINS
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '12 - Segurança', 'Logins sysadmin', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 5 THEN 'ATENÇÃO' ELSE 'INFO' END,
       'Quantidade alta não implica vulnerabilidade; revisar necessidade e governança.'
FROM sys.server_principals sp
JOIN sys.server_role_members srm ON sp.principal_id = srm.member_principal_id
JOIN sys.server_principals rolep ON rolep.principal_id = srm.role_principal_id
WHERE rolep.name = 'sysadmin';

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '12 - Segurança', 'Login sa habilitado', CASE WHEN is_disabled = 1 THEN 'NÃO' ELSE 'SIM' END,
       CASE WHEN is_disabled = 1 THEN 'OK' ELSE 'ATENÇÃO' END,
       'Validar conforme política de segurança do cliente.'
FROM sys.server_principals
WHERE sid = 0x01;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'12 - Segurança', N'Logins SQL habilitados sem política de senha', CONVERT(nvarchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'CHECK_POLICY = OFF: sem complexidade/bloqueio de senha do Windows.'
FROM sys.sql_logins
WHERE is_disabled = 0 AND is_policy_checked = 0 AND name NOT LIKE N'##%';

/* ============================================================================
   13 - CONFIGURAÇÕES DA INSTÂNCIA
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'13 - Configuração', N'max server memory (MB)', CONVERT(nvarchar(30), value_in_use),
       CASE WHEN CONVERT(bigint, value_in_use) >= 2147483647 THEN N'ATENÇÃO' ELSE N'INFO' END,
       CASE WHEN CONVERT(bigint, value_in_use) >= 2147483647
            THEN N'Valor padrão (não configurado): o SQL Server pode consumir praticamente toda a memória do SO.'
            ELSE N'Memória máxima configurada para o SQL Server.' END
FROM sys.configurations WHERE name = 'max server memory (MB)'
UNION ALL
SELECT '13 - Configuração', 'min server memory (MB)', CONVERT(varchar(30), value_in_use), 'INFO', 'Memória mínima configurada.'
FROM sys.configurations WHERE name = 'min server memory (MB)'
UNION ALL
SELECT N'13 - Configuração', N'cost threshold for parallelism', CONVERT(nvarchar(30), value_in_use),
       CASE WHEN CONVERT(int, value_in_use) = 5 THEN N'ATENÇÃO' ELSE N'INFO' END,
       CASE WHEN CONVERT(int, value_in_use) = 5
            THEN N'Valor padrão (5) costuma ser baixo para cargas atuais; avaliar ajuste (ex.: 30-50) com base no workload.'
            ELSE N'Threshold configurado.' END
FROM sys.configurations WHERE name = 'cost threshold for parallelism'
UNION ALL
SELECT N'13 - Configuração', N'max degree of parallelism', CONVERT(nvarchar(30), value_in_use),
       CASE WHEN CONVERT(int, value_in_use) = 0 AND @Dashboard_LogicalCPU > 8 THEN N'ATENÇÃO' ELSE N'INFO' END,
       CASE WHEN CONVERT(int, value_in_use) = 0 AND @Dashboard_LogicalCPU > 8
            THEN N'MAXDOP = 0 com mais de 8 CPUs lógicas; revisar conforme recomendação Microsoft por NUMA.'
            ELSE N'MAXDOP configurado.' END
FROM sys.configurations WHERE name = 'max degree of parallelism'
UNION ALL
SELECT '13 - Configuração', 'backup compression default', CONVERT(varchar(30), value_in_use), 'INFO', 'Compressão padrão de backup.'
FROM sys.configurations WHERE name = 'backup compression default'
UNION ALL
SELECT '13 - Configuração', 'optimize for ad hoc workloads', CONVERT(varchar(30), value_in_use), 'INFO', 'Configuração de planos ad hoc.'
FROM sys.configurations WHERE name = 'optimize for ad hoc workloads'
UNION ALL
SELECT '13 - Configuração', 'remote admin connections', CONVERT(varchar(30), value_in_use), 'INFO', 'DAC remoto.'
FROM sys.configurations WHERE name = 'remote admin connections'
UNION ALL
SELECT N'13 - Configuração', N'xp_cmdshell', CONVERT(nvarchar(30), value_in_use),
       CASE WHEN CONVERT(int, value_in_use) = 1 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'Habilitado amplia a superfície de ataque; validar necessidade.'
FROM sys.configurations WHERE name = 'xp_cmdshell';

/* ============================================================================
   14 - MEMÓRIA / CPU / PLE
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '14 - Memória', 'SQL Server Memory Usage (MB)', CONVERT(varchar(30), CONVERT(decimal(18,2), physical_memory_in_use_kb/1024.0)),
       CASE WHEN process_physical_memory_low = 1 THEN 'ATENÇÃO' ELSE 'INFO' END,
       'process_physical_memory_low deve normalmente estar em 0.'
FROM sys.dm_os_process_memory;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '14 - Memória', 'Locked Pages Allocation (MB)', CONVERT(varchar(30), CONVERT(decimal(18,2), locked_page_allocations_kb/1024.0)), 'INFO', 'Indica uso de locked pages pelo SQL Server.'
FROM sys.dm_os_process_memory;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '14 - Memória', 'Estado de memória do SO', system_memory_state_desc,
       CASE WHEN system_memory_state_desc LIKE '%low%' OR system_memory_state_desc LIKE '%running low%' THEN 'ATENÇÃO' ELSE 'INFO' END,
       'Indicador de pressão de memória no sistema operacional.'
FROM sys.dm_os_sys_memory;

INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '14 - Memória', 'PLE médio', CONVERT(varchar(30), CONVERT(decimal(18,2), AVG(CONVERT(float,cntr_value)))), 'INFO', 'Usar tendência e contexto; não utilizar valor isolado como diagnóstico.'
FROM sys.dm_os_performance_counters
WHERE object_name LIKE '%Buffer Node%' AND counter_name = 'Page life expectancy';

/* ============================================================================
   15 - I/O / WARNINGS
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '15 - I/O', 'Arquivos com latência média > 30 ms', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'ATENÇÃO' ELSE 'OK' END,
       'Indicador inicial; investigar por workload, storage e janela de tempo.'
FROM
(
    SELECT mf.database_id, mf.file_id,
           CASE WHEN vfs.num_of_reads + vfs.num_of_writes = 0 THEN 0
                ELSE (vfs.io_stall_read_ms + vfs.io_stall_write_ms) * 1.0 / (vfs.num_of_reads + vfs.num_of_writes) END AS AvgLatency
    FROM sys.dm_io_virtual_file_stats(NULL,NULL) vfs
    JOIN sys.master_files mf ON mf.database_id=vfs.database_id AND mf.file_id=vfs.file_id
) X
WHERE AvgLatency > 30;

/* ============================================================================
   16 - SUSPECT PAGES / INTEGRIDADE
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT '16 - Integridade', 'Suspect Pages ativas', CONVERT(varchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN 'CRÍTICO' ELSE 'OK' END,
       'Registros na msdb.dbo.suspect_pages devem ser investigados.'
FROM msdb.dbo.suspect_pages
WHERE event_type IN (1,2,3);

/* ============================================================================
   17 - LOG DE TRANSAÇÕES (V9)
============================================================================ */
INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
SELECT N'17 - Log', N'Logs com uso acima de 80%', CONVERT(nvarchar(30), COUNT(*)),
       CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
       N'Cruzar com log_reuse_wait_desc no detalhamento de databases.'
FROM #LogSpace
WHERE LogSpaceUsedPct > 80;

BEGIN TRY
    SELECT d.name AS DatabaseName,
           li.file_id,
           li.VLFCount
    INTO #D17_VLF
    FROM sys.databases d
    CROSS APPLY
    (
        SELECT file_id, COUNT(*) AS VLFCount
        FROM sys.dm_db_log_info(d.database_id)
        GROUP BY file_id
    ) li
    WHERE d.state_desc = 'ONLINE' AND HAS_DBACCESS(d.name) = 1;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'17 - Log', N'Arquivos de log com mais de 1.000 VLFs', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Excesso de VLFs aumenta o tempo de recovery, restore e sincronização de AG.'
    FROM #D17_VLF
    WHERE VLFCount > 1000;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'17 - VLF', ERROR_MESSAGE());
END CATCH;

/* ============================================================================
   18 a 24 - COLETA COMPLEMENTAR (V10)
   Segurança avançada, objetos, fragmentação, missing indexes, top queries,
   sessões, histórico de backup e agendamentos de jobs.
   Os dados ficam em #D19..#D24 e são registrados em #Secoes após a seção 18.
============================================================================ */

/* ---------- 19 - Segurança complementar ---------- */
BEGIN TRY
    SELECT
        pr.name                                   AS PermPrincipal,
        pr.type_desc                              AS PermPrincipalType,
        pe.permission_name                        AS PermName,
        pe.state_desc                             AS PermState,
        pe.class_desc                             AS PermClass,
        CONVERT(nvarchar(256), CASE pe.class
            WHEN 105 THEN (SELECT e.name FROM sys.endpoints e WHERE e.endpoint_id = pe.major_id)
            WHEN 101 THEN SUSER_NAME(pe.major_id)
            ELSE N''
        END)                                      AS PermObject,
        ISNULL(gr.name, N'')                      AS PermGrantor
    INTO #D19_Permissoes
    FROM sys.server_permissions pe
    JOIN sys.server_principals pr      ON pr.principal_id = pe.grantee_principal_id
    LEFT JOIN sys.server_principals gr ON gr.principal_id = pe.grantor_principal_id
    WHERE pr.name NOT LIKE N'##%'
      AND NOT (pe.permission_name = N'CONNECT SQL' AND pe.state = 'G')
      AND NOT (pr.name = N'public' AND pe.permission_name IN (N'VIEW ANY DATABASE', N'CONNECT'));

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'18 - Segurança (V10)', N'Permissões de servidor de alto privilégio concedidas diretamente', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'CONTROL SERVER, ALTER ANY LOGIN, IMPERSONATE ANY LOGIN etc. concedidas fora das server roles equivalem (ou quase) a sysadmin.'
    FROM #D19_Permissoes
    WHERE PermState LIKE N'GRANT%'
      AND PermName IN (N'CONTROL SERVER', N'ALTER ANY LOGIN', N'IMPERSONATE ANY LOGIN', N'ALTER ANY SERVER ROLE',
                       N'UNSAFE ASSEMBLY', N'ALTER SETTINGS', N'ALTER ANY DATABASE', N'SHUTDOWN',
                       N'ALTER ANY CREDENTIAL', N'EXTERNAL ACCESS ASSEMBLY');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'19 - Permissões de servidor', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT
        r.name                         AS SrvRoleName,
        m.name                         AS SrvRoleMember,
        m.type_desc                    AS SrvRoleMemberType,
        m.is_disabled                  AS SrvRoleMemberDisabled
    INTO #D19_ServerRoles
    FROM sys.server_role_members rm
    JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id
    JOIN sys.server_principals m ON m.principal_id = rm.member_principal_id
    WHERE m.name NOT LIKE N'##%';

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'18 - Segurança (V10)', N'Membros ativos de securityadmin', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'securityadmin pode conceder qualquer permissão, inclusive sysadmin a si mesmo. Tratar como sysadmin.'
    FROM #D19_ServerRoles
    WHERE SrvRoleName = N'securityadmin' AND SrvRoleMemberDisabled = 0;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'19 - Server roles', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT
        l.name                                                          AS PolLogin,
        l.is_disabled                                                   AS PolDisabled,
        l.is_policy_checked                                             AS PolPolicyChecked,
        l.is_expiration_checked                                         AS PolExpirationChecked,
        CONVERT(datetime, LOGINPROPERTY(l.name, 'PasswordLastSetTime')) AS PolPasswordLastSet,
        CONVERT(bit, LOGINPROPERTY(l.name, 'IsLocked'))                 AS PolIsLocked,
        CONVERT(bit, LOGINPROPERTY(l.name, 'IsExpired'))                AS PolIsExpired,
        CONVERT(int, LOGINPROPERTY(l.name, 'BadPasswordCount'))         AS PolBadPwdCount,
        CONVERT(bit, CASE WHEN PWDCOMPARE(N'', l.password_hash) = 1 THEN 1 ELSE 0 END)     AS PolBlankPassword,
        CONVERT(bit, CASE WHEN PWDCOMPARE(l.name, l.password_hash) = 1 THEN 1 ELSE 0 END)  AS PolPasswordEqualsLogin,
        l.default_database_name                                         AS PolDefaultDb
    INTO #D19_PoliticaSenha
    FROM sys.sql_logins l
    WHERE l.name NOT LIKE N'##%';

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'18 - Segurança (V10)', N'Logins SQL habilitados com senha em branco', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
           N'Verificado com PWDCOMPARE. Trocar a senha imediatamente.'
    FROM #D19_PoliticaSenha WHERE PolDisabled = 0 AND PolBlankPassword = 1;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'18 - Segurança (V10)', N'Logins SQL habilitados com senha igual ao nome do login', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
           N'Verificado com PWDCOMPARE. Primeira tentativa de qualquer ataque de força bruta.'
    FROM #D19_PoliticaSenha WHERE PolDisabled = 0 AND PolPasswordEqualsLogin = 1;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'18 - Segurança (V10)', N'Logins SQL habilitados sem expiração de senha', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'CHECK_EXPIRATION = OFF. Aceitável para contas de aplicação com rotação controlada; para pessoas, ativar.'
    FROM #D19_PoliticaSenha WHERE PolDisabled = 0 AND PolExpirationChecked = 0;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'18 - Segurança (V10)', N'Logins SQL bloqueados por tentativas de senha', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Pode indicar aplicação com senha desatualizada ou tentativa de força bruta.'
    FROM #D19_PoliticaSenha WHERE PolIsLocked = 1;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'19 - Política de senha', ERROR_MESSAGE());
END CATCH;

/* ---------- 20 - Objetos e fragmentação (loop por database) ---------- */
CREATE TABLE #D20_Objetos (
    ObjDatabase sysname, ObjTabelas int, ObjViews int, ObjProcedures int, ObjFunctions int,
    ObjTriggers int, ObjIndices int, ObjSinonimos int, ObjSchemas int);

CREATE TABLE #V10_FragBase (
    FragDatabase sysname, FragSchema sysname, FragTable sysname, FragIndex sysname NULL,
    FragIndexType nvarchar(60), FragIndexId int, FragPct decimal(5,1), FragPages bigint);

DECLARE @V10_db sysname, @V10_sql nvarchar(max);
DECLARE curV10 CURSOR LOCAL FAST_FORWARD FOR
    SELECT name FROM sys.databases
    WHERE state = 0 AND database_id <> 2 AND source_database_id IS NULL AND HAS_DBACCESS(name) = 1
    ORDER BY name;
OPEN curV10;
FETCH NEXT FROM curV10 INTO @V10_db;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @V10_sql = N'USE ' + QUOTENAME(@V10_db) + N';
    INSERT #D20_Objetos
    SELECT DB_NAME(),
        SUM(CASE WHEN o.type = ''U'' THEN 1 ELSE 0 END),
        SUM(CASE WHEN o.type = ''V'' THEN 1 ELSE 0 END),
        SUM(CASE WHEN o.type IN (''P'',''PC'') THEN 1 ELSE 0 END),
        SUM(CASE WHEN o.type IN (''FN'',''IF'',''TF'',''FS'',''FT'') THEN 1 ELSE 0 END),
        SUM(CASE WHEN o.type = ''TR'' THEN 1 ELSE 0 END),
        (SELECT COUNT(*) FROM sys.indexes i JOIN sys.objects x ON x.object_id = i.object_id
          WHERE x.is_ms_shipped = 0 AND i.index_id > 0 AND x.type IN (''U'',''V'')),
        SUM(CASE WHEN o.type = ''SN'' THEN 1 ELSE 0 END),
        (SELECT COUNT(*) FROM sys.schemas s WHERE s.schema_id BETWEEN 5 AND 16383)
    FROM sys.objects o
    WHERE o.is_ms_shipped = 0;';
    BEGIN TRY
        EXEC sys.sp_executesql @V10_sql;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'20 - Objetos (' + @V10_db + N')', ERROR_MESSAGE());
    END CATCH;

    IF @ColetarFragmentacao = 1 AND DB_ID(@V10_db) > 4
    BEGIN
        SET @V10_sql = N'USE ' + QUOTENAME(@V10_db) + N';
        INSERT #V10_FragBase
        SELECT DB_NAME(), s.name, o.name, i.name, i.type_desc, i.index_id,
               CONVERT(decimal(5,1), ps.avg_fragmentation_in_percent), ps.page_count
        FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, ''LIMITED'') ps
        JOIN sys.indexes i ON i.object_id = ps.object_id AND i.index_id = ps.index_id
        JOIN sys.objects o ON o.object_id = i.object_id
        JOIN sys.schemas s ON s.schema_id = o.schema_id
        WHERE ps.index_id > 0
          AND ps.alloc_unit_type_desc = ''IN_ROW_DATA''
          AND ps.page_count >= @minp
          AND o.is_ms_shipped = 0;';
        BEGIN TRY
            EXEC sys.sp_executesql @V10_sql, N'@minp int', @minp = @FragMinPaginas;
        END TRY
        BEGIN CATCH
            INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'20 - Fragmentação (' + @V10_db + N')', ERROR_MESSAGE());
        END CATCH;
    END;

    FETCH NEXT FROM curV10 INTO @V10_db;
END;
CLOSE curV10;
DEALLOCATE curV10;

SELECT TOP (500)
    FragDatabase, FragSchema, FragTable, ISNULL(FragIndex, N'(heap)') AS FragIndex,
    FragIndexType, FragIndexId, FragPct, FragPages,
    CONVERT(nvarchar(20), CASE WHEN FragPct >= 70 THEN N'REBUILD Crítico' ELSE N'REBUILD' END) AS FragAction
INTO #D20_Fragmentacao
FROM #V10_FragBase
WHERE FragPct >= 30
ORDER BY FragPct DESC, FragPages DESC;

SELECT
    FragDatabase                                                    AS FragResDatabase,
    SUM(CASE WHEN FragPct >= 70 THEN 1 ELSE 0 END)                  AS FragResCriticos,
    SUM(CASE WHEN FragPct >= 30 AND FragPct < 70 THEN 1 ELSE 0 END) AS FragResModerados,
    SUM(CASE WHEN FragPct < 30 THEN 1 ELSE 0 END)                   AS FragResOk,
    COUNT(*)                                                        AS FragResTotal,
    @FragMinPaginas                                                 AS FragResMinPaginas
INTO #D20_FragResumo
FROM #V10_FragBase
GROUP BY FragDatabase;

IF @ColetarFragmentacao = 1
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'19 - Índices (V10)', N'Índices com fragmentação >= 70%', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Somente índices com ' + CONVERT(nvarchar(20), @FragMinPaginas) + N' páginas ou mais (modo LIMITED). Avaliar REBUILD/rotina de manutenção.'
    FROM #V10_FragBase WHERE FragPct >= 70;
ELSE
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    VALUES (N'19 - Índices (V10)', N'Índices com fragmentação >= 70%', N'Não coletado', N'INFO', N'Coleta de fragmentação desligada (@ColetarFragmentacao = 0).');

/* ---------- 20 - Missing indexes ---------- */
BEGIN TRY
    SELECT TOP (25)
        DB_NAME(mid.database_id)                                        AS MIDatabase,
        mid.statement                                                   AS MITable,
        ISNULL(mid.equality_columns, N'')                               AS MIEquality,
        ISNULL(mid.inequality_columns, N'')                             AS MIInequality,
        ISNULL(mid.included_columns, N'')                               AS MIIncluded,
        migs.user_seeks                                                 AS MIUserSeeks,
        migs.user_scans                                                 AS MIUserScans,
        CONVERT(decimal(5,1), migs.avg_user_impact)                     AS MIAvgImpactPct,
        CONVERT(decimal(18,1), migs.avg_total_user_cost * (migs.avg_user_impact / 100.0)
                 * (migs.user_seeks + migs.user_scans))                 AS MIImprovement,
        migs.last_user_seek                                             AS MILastSeek,
        CONVERT(nvarchar(4000), N'CREATE NONCLUSTERED INDEX [IX_' + ISNULL(OBJECT_NAME(mid.object_id, mid.database_id), N'tabela')
            + N'_sug' + CONVERT(nvarchar(12), mig.index_group_handle) + N'] ON ' + mid.statement + N' ('
            + ISNULL(mid.equality_columns, N'')
            + CASE WHEN mid.equality_columns IS NOT NULL AND mid.inequality_columns IS NOT NULL THEN N', ' ELSE N'' END
            + ISNULL(mid.inequality_columns, N'') + N')'
            + ISNULL(N' INCLUDE (' + mid.included_columns + N')', N'') + N';') AS MICreateStatement
    INTO #D20_MissingIndexes
    FROM sys.dm_db_missing_index_group_stats migs
    JOIN sys.dm_db_missing_index_groups  mig ON migs.group_handle = mig.index_group_handle
    JOIN sys.dm_db_missing_index_details mid ON mig.index_handle  = mid.index_handle
    WHERE mid.database_id > 4
    ORDER BY migs.avg_total_user_cost * (migs.avg_user_impact / 100.0) * (migs.user_seeks + migs.user_scans) DESC;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'19 - Índices (V10)', N'Índices sugeridos pelo engine (impacto >= 80%)', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Top 25 do sys.dm_db_missing_index_*. Consolidar sugestões sobrepostas e avaliar o custo de escrita antes de criar.'
    FROM #D20_MissingIndexes WHERE MIAvgImpactPct >= 80 AND MIUserSeeks + MIUserScans >= 100;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'20 - Missing indexes', ERROR_MESSAGE());
END CATCH;

/* ---------- 21 - Top queries (plan cache) ---------- */
BEGIN TRY
    SELECT TOP (20)
        ISNULL(DB_NAME(CONVERT(int, pa.value)), N'-')                                  AS TQDatabase,
        CONVERT(decimal(18,2), qs.total_elapsed_time / 1000.0 / qs.execution_count)    AS TQAvgDurationMs,
        qs.execution_count                                                             AS TQExecutions,
        CONVERT(decimal(18,2), qs.total_worker_time / 1000.0 / qs.execution_count)     AS TQAvgCpuMs,
        qs.total_logical_reads  / qs.execution_count                                   AS TQAvgLogicalReads,
        qs.total_physical_reads / qs.execution_count                                   AS TQAvgPhysicalReads,
        qs.creation_time                                                               AS TQPlanCreated,
        qs.last_execution_time                                                         AS TQLastExecution,
        CONVERT(nvarchar(2000), REPLACE(REPLACE(REPLACE(LEFT(SUBSTRING(st.text, (qs.statement_start_offset / 2) + 1,
            ((CASE qs.statement_end_offset WHEN -1 THEN DATALENGTH(st.text) ELSE qs.statement_end_offset END
              - qs.statement_start_offset) / 2) + 1), 2000), NCHAR(13), N' '), NCHAR(10), N' '), NCHAR(9), N' ')) AS TQSqlText
    INTO #D21_TopDuracao
    FROM sys.dm_exec_query_stats qs
    CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) st
    OUTER APPLY (SELECT a.value FROM sys.dm_exec_plan_attributes(qs.plan_handle) a WHERE a.attribute = N'dbid') pa
    WHERE qs.execution_count >= 2
    ORDER BY qs.total_elapsed_time / qs.execution_count DESC;

    SELECT TOP (20)
        ISNULL(DB_NAME(CONVERT(int, pa.value)), N'-')                                  AS TCDatabase,
        CONVERT(decimal(18,2), qs.total_worker_time  / 1000.0)                         AS TCTotalCpuMs,
        CONVERT(decimal(18,2), qs.total_worker_time  / 1000.0 / qs.execution_count)    AS TCAvgCpuMs,
        qs.execution_count                                                             AS TCExecutions,
        CONVERT(decimal(18,2), qs.total_elapsed_time / 1000.0)                         AS TCTotalDurationMs,
        qs.total_logical_reads                                                         AS TCTotalLogicalReads,
        qs.last_execution_time                                                         AS TCLastExecution,
        CONVERT(nvarchar(2000), REPLACE(REPLACE(REPLACE(LEFT(SUBSTRING(st.text, (qs.statement_start_offset / 2) + 1,
            ((CASE qs.statement_end_offset WHEN -1 THEN DATALENGTH(st.text) ELSE qs.statement_end_offset END
              - qs.statement_start_offset) / 2) + 1), 2000), NCHAR(13), N' '), NCHAR(10), N' '), NCHAR(9), N' ')) AS TCSqlText
    INTO #D21_TopCPU
    FROM sys.dm_exec_query_stats qs
    CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) st
    OUTER APPLY (SELECT a.value FROM sys.dm_exec_plan_attributes(qs.plan_handle) a WHERE a.attribute = N'dbid') pa
    ORDER BY qs.total_worker_time DESC;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'20 - Queries (V10)', N'Queries com duração média > 5 s (top 20 do plan cache)', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Baseado em sys.dm_exec_query_stats (somente planos em cache, ao menos 2 execuções). Revisar planos e índices.'
    FROM #D21_TopDuracao WHERE TQAvgDurationMs > 5000;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'21 - Top queries', ERROR_MESSAGE());
END CATCH;

/* ---------- 22 - Sessões de usuário ---------- */
BEGIN TRY
    SELECT
        s.session_id                                                    AS SessSessionId,
        s.login_name                                                    AS SessLogin,
        ISNULL(s.host_name, N'')                                        AS SessHost,
        CONVERT(nvarchar(128), ISNULL(s.program_name, N''))             AS SessProgram,
        ISNULL(DB_NAME(s.database_id), N'')                             AS SessDatabase,
        CONVERT(nvarchar(30), ISNULL(r.status, s.status))               AS SessStatus,
        ISNULL(r.blocking_session_id, 0)                                AS SessBlockedBy,
        ISNULL(r.wait_type, N'')                                        AS SessWaitType,
        ISNULL(r.wait_time, 0)                                          AS SessWaitMs,
        s.cpu_time                                                      AS SessCpuMs,
        s.reads                                                         AS SessReads,
        s.writes                                                        AS SessWrites,
        s.logical_reads                                                 AS SessLogicalReads,
        s.open_transaction_count                                        AS SessOpenTran,
        DATEDIFF(SECOND, ISNULL(r.start_time, s.last_request_start_time), GETDATE()) AS SessElapsedSec,
        s.login_time                                                    AS SessLoginTime,
        s.last_request_end_time                                         AS SessLastRequestEnd,
        CONVERT(nvarchar(1000), REPLACE(REPLACE(REPLACE(LEFT(ISNULL(st.text, N''), 1000), NCHAR(13), N' '), NCHAR(10), N' '), NCHAR(9), N' ')) AS SessSqlText
    INTO #D22_Sessoes
    FROM sys.dm_exec_sessions s
    LEFT JOIN sys.dm_exec_requests r ON r.session_id = s.session_id
    OUTER APPLY (SELECT TOP (1) c.most_recent_sql_handle FROM sys.dm_exec_connections c WHERE c.session_id = s.session_id) c
    OUTER APPLY sys.dm_exec_sql_text(COALESCE(r.sql_handle, c.most_recent_sql_handle)) st
    WHERE s.is_user_process = 1
      AND s.session_id <> @@SPID;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'21 - Sessões (V10)', N'Sessões bloqueadas no momento da coleta', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Fotografia do momento da coleta. Recorrência deve ser investigada com blocked process report / sp_WhoIsActive.'
    FROM #D22_Sessoes WHERE SessBlockedBy <> 0;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'21 - Sessões (V10)', N'Sessões ociosas com transação aberta', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Transações esquecidas seguram locks e impedem a reutilização do log.'
    FROM #D22_Sessoes WHERE SessOpenTran > 0 AND SessStatus = N'sleeping';
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'22 - Sessões', ERROR_MESSAGE());
END CATCH;

/* ---------- 23 - Histórico de backups (7 dias, msdb local) ---------- */
BEGIN TRY
    SELECT
        bs.database_name                                                            AS BkDatabase,
        CONVERT(nvarchar(20), CASE bs.type WHEN 'D' THEN N'Full' WHEN 'I' THEN N'Diferencial' ELSE bs.type END) AS BkType,
        bs.backup_start_date                                                        AS BkStart,
        bs.backup_finish_date                                                       AS BkFinish,
        DATEDIFF(SECOND, bs.backup_start_date, bs.backup_finish_date)               AS BkDurationSec,
        CONVERT(decimal(18,2), bs.backup_size / 1048576.0)                          AS BkSizeMB,
        CONVERT(decimal(18,2), bs.compressed_backup_size / 1048576.0)               AS BkCompressedMB,
        CONVERT(decimal(9,2), bs.backup_size / NULLIF(bs.compressed_backup_size, 0)) AS BkCompressionRatio,
        bs.is_copy_only                                                             AS BkCopyOnly,
        bs.is_snapshot                                                              AS BkSnapshot,
        bs.server_name                                                              AS BkServer,
        bs.user_name                                                                AS BkUser,
        ISNULL(mf.device_type, 0)                                                   AS BkDeviceType,
        CONVERT(nvarchar(400), ISNULL(mf.physical_device_name, N''))                AS BkDevice
    INTO #D23_BackupHist
    FROM msdb.dbo.backupset bs
    LEFT JOIN msdb.dbo.backupmediafamily mf
           ON mf.media_set_id = bs.media_set_id AND mf.family_sequence_number = 1
    WHERE bs.type IN ('D', 'I')
      AND bs.backup_start_date >= DATEADD(DAY, -7, GETDATE());

    SELECT
        bs.database_name                                                    AS LogBkDatabase,
        CONVERT(char(10), bs.backup_start_date, 120)                        AS LogBkDay,
        COUNT(*)                                                            AS LogBkCount,
        CONVERT(decimal(18,2), SUM(bs.backup_size) / 1048576.0)             AS LogBkSizeMB,
        CONVERT(decimal(18,2), SUM(bs.compressed_backup_size) / 1048576.0)  AS LogBkCompressedMB,
        MIN(bs.backup_start_date)                                           AS LogBkFirst,
        MAX(bs.backup_finish_date)                                          AS LogBkLast
    INTO #D23_BackupLog
    FROM msdb.dbo.backupset bs
    WHERE bs.type = 'L'
      AND bs.backup_start_date >= DATEADD(DAY, -7, GETDATE())
    GROUP BY bs.database_name, CONVERT(char(10), bs.backup_start_date, 120);

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'22 - Histórico de backup (V10)', N'Backups Full/Diff (7 dias) gravados no mesmo volume dos dados', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
           N'Se o disco falhar, perdem-se os dados e os backups juntos. Gravar em outro volume/servidor e manter cópia fora do site.'
    FROM #D23_BackupHist h
    WHERE h.BkDeviceType = 2 AND h.BkSnapshot = 0
      AND h.BkDevice LIKE N'_:\%'
      AND LEFT(h.BkDevice, 3) IN (SELECT DISTINCT LEFT(physical_name, 3) FROM sys.master_files WHERE type = 0);

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'22 - Histórico de backup (V10)', N'Backups nativos (7 dias) sem compressão', CONVERT(nvarchar(30), COUNT(*)),
           CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
           N'Taxa de compressão ~1,0. Usar WITH COMPRESSION ou ativar backup compression default.'
    FROM #D23_BackupHist
    WHERE BkDeviceType <> 7 AND BkSnapshot = 0 AND ISNULL(BkCompressionRatio, 1) <= 1.05;

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'22 - Histórico de backup (V10)', N'Backups por ferramenta externa (VDI/snapshot) em 7 dias', CONVERT(nvarchar(30), COUNT(*)),
           N'INFO',
           N'Ex.: Veeam, Commvault. Confirmar retenção e testes de restore na própria ferramenta. Em Always On, o msdb mostra só os backups desta réplica.'
    FROM #D23_BackupHist
    WHERE BkDeviceType = 7 OR BkSnapshot = 1;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'23 - Histórico de backup', ERROR_MESSAGE());
END CATCH;

/* ---------- 24 - Agendamentos dos jobs ---------- */
BEGIN TRY
    SELECT
        j.name                                                       AS SchJobName,
        j.enabled                                                    AS SchJobEnabled,
        ISNULL(ss.name, N'(sem agendamento)')                        AS SchScheduleName,
        CONVERT(bit, ISNULL(ss.enabled, 0))                          AS SchEnabled,
        CONVERT(nvarchar(40), CASE ss.freq_type
            WHEN 1   THEN N'Uma vez'
            WHEN 4   THEN CASE WHEN ss.freq_interval > 1 THEN N'A cada ' + CONVERT(nvarchar(5), ss.freq_interval) + N' dias' ELSE N'Diário' END
            WHEN 8   THEN N'Semanal'
            WHEN 16  THEN N'Mensal'
            WHEN 32  THEN N'Mensal relativo'
            WHEN 64  THEN N'Ao iniciar o SQL Agent'
            WHEN 128 THEN N'Quando a CPU estiver ociosa'
            ELSE N'-'
        END)                                                         AS SchFrequency,
        CONVERT(nvarchar(60), CASE
            WHEN ss.freq_type = 8 THEN ISNULL(STUFF(
                  CASE WHEN ss.freq_interval &  1 =  1 THEN N',Dom' ELSE N'' END
                + CASE WHEN ss.freq_interval &  2 =  2 THEN N',Seg' ELSE N'' END
                + CASE WHEN ss.freq_interval &  4 =  4 THEN N',Ter' ELSE N'' END
                + CASE WHEN ss.freq_interval &  8 =  8 THEN N',Qua' ELSE N'' END
                + CASE WHEN ss.freq_interval & 16 = 16 THEN N',Qui' ELSE N'' END
                + CASE WHEN ss.freq_interval & 32 = 32 THEN N',Sex' ELSE N'' END
                + CASE WHEN ss.freq_interval & 64 = 64 THEN N',Sáb' ELSE N'' END, 1, 1, N''), N'')
            WHEN ss.freq_type = 16 THEN N'Dia ' + CONVERT(nvarchar(5), ss.freq_interval)
            ELSE N''
        END)                                                         AS SchDays,
        CONVERT(nvarchar(60), CASE ss.freq_subday_type
            WHEN 1 THEN N'Uma vez às ' + STUFF(STUFF(RIGHT(N'000000' + CONVERT(nvarchar(6), ss.active_start_time), 6), 5, 0, N':'), 3, 0, N':')
            WHEN 2 THEN N'A cada ' + CONVERT(nvarchar(5), ss.freq_subday_interval) + N' segundo(s)'
            WHEN 4 THEN N'A cada ' + CONVERT(nvarchar(5), ss.freq_subday_interval) + N' minuto(s)'
            WHEN 8 THEN N'A cada ' + CONVERT(nvarchar(5), ss.freq_subday_interval) + N' hora(s)'
            ELSE N''
        END)                                                         AS SchSubday,
        CONVERT(nvarchar(8), CASE WHEN ss.schedule_id IS NULL THEN N'' ELSE STUFF(STUFF(RIGHT(N'000000' + CONVERT(nvarchar(6), ss.active_start_time), 6), 5, 0, N':'), 3, 0, N':') END) AS SchStartTime,
        CONVERT(nvarchar(8), CASE WHEN ss.schedule_id IS NULL THEN N'' ELSE STUFF(STUFF(RIGHT(N'000000' + CONVERT(nvarchar(6), ss.active_end_time),   6), 5, 0, N':'), 3, 0, N':') END) AS SchEndTime,
        CASE WHEN js.next_run_date > 0 THEN msdb.dbo.agent_datetime(js.next_run_date, js.next_run_time) END AS SchNextRun
    INTO #D24_JobAgenda
    FROM msdb.dbo.sysjobs j
    JOIN msdb.dbo.syscategories cat       ON cat.category_id = j.category_id
    LEFT JOIN msdb.dbo.sysjobschedules js ON js.job_id = j.job_id
    LEFT JOIN msdb.dbo.sysschedules ss    ON ss.schedule_id = js.schedule_id
    WHERE cat.name <> N'Report Server';

    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    SELECT N'11 - SQL Agent', N'Jobs habilitados sem agendamento ativo (V10)', CONVERT(nvarchar(30), COUNT(*)),
           N'INFO',
           N'Só executam manualmente ou por outro job. Confirmar se é intencional.'
    FROM (SELECT SchJobName FROM #D24_JobAgenda WHERE SchJobEnabled = 1
          GROUP BY SchJobName HAVING MAX(CONVERT(int, SchEnabled)) = 0) x;
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'24 - Agendamentos', ERROR_MESSAGE());
END CATCH;

/* ============================================================================
   25 - LOG SHIPPING (V11)
   Configuração das databases primárias e secundárias, jobs LSBackup/LSCopy/
   LSRestore/LSAlert, atraso do último backup/restore x limites do monitor e
   erros registrados pelo log shipping. Fonte: msdb.dbo.log_shipping_*.
   Os dados ficam em #D25_* e são registrados em #Secoes (grupo 25).
============================================================================ */
CREATE TABLE #V11_LSJob (
    job_id uniqueidentifier NOT NULL, JobName sysname NOT NULL, JobEnabled bit NOT NULL,
    JobLastRun datetime NULL, JobLastOutcome nvarchar(20) NOT NULL, JobNextRun datetime NULL);

CREATE TABLE #D25_LSPrimario (
    LSPDatabase sysname NOT NULL, LSPStatus nvarchar(20) NOT NULL, LSPUltimoBackup datetime NULL,
    LSPMinDesdeBackup int NULL, LSPLimiteMin int NULL, LSPAlertaHabilitado bit NULL,
    LSPSecundarios nvarchar(2000) NULL, LSPMonitor nvarchar(300) NULL, LSPJobBackup sysname NULL,
    LSPCompressao nvarchar(30) NULL, LSPRetencaoMin int NULL, LSPBackupDir nvarchar(500) NULL,
    LSPBackupShare nvarchar(500) NULL, LSPUltimoArquivo nvarchar(500) NULL);

CREATE TABLE #D25_LSSecundario (
    LSSDatabase sysname NOT NULL, LSSStatus nvarchar(20) NOT NULL, LSSServidorPrimario nvarchar(300) NULL,
    LSSDatabasePrimaria sysname NULL, LSSUltimaCopia datetime NULL, LSSUltimoRestore datetime NULL,
    LSSMinDesdeRestore int NULL, LSSLatenciaMin int NULL, LSSLimiteMin int NULL, LSSAtrasoRestoreMin int NULL,
    LSSAlertaHabilitado bit NULL, LSSModoRestore nvarchar(30) NULL, LSSDesconectaUsuarios bit NULL,
    LSSMonitor nvarchar(300) NULL, LSSJobCopia sysname NULL, LSSJobRestore sysname NULL,
    LSSRetencaoMin int NULL, LSSPastaOrigem nvarchar(500) NULL, LSSPastaDestino nvarchar(500) NULL,
    LSSUltimoArquivo nvarchar(500) NULL);

CREATE TABLE #D25_LSMonitor (
    LSMPapel nvarchar(20) NOT NULL, LSMServidor nvarchar(300) NULL, LSMDatabase sysname NULL,
    LSMStatus nvarchar(20) NOT NULL, LSMParceiro nvarchar(600) NULL, LSMUltimaOperacao datetime NULL,
    LSMMinDesde int NULL, LSMLatenciaMin int NULL, LSMLimiteMin int NULL, LSMAlertaHabilitado bit NULL);

CREATE TABLE #D25_LSJobs (
    LSJFuncao nvarchar(20) NOT NULL, LSJJob sysname NOT NULL, LSJDatabase nvarchar(300) NULL,
    LSJHabilitado bit NOT NULL, LSJUltimaExecucao datetime NULL, LSJResultado nvarchar(20) NOT NULL,
    LSJProximaExecucao datetime NULL, LSJStatus nvarchar(20) NOT NULL);

CREATE TABLE #D25_LSErros (
    LSEData datetime NULL, LSEAgente nvarchar(20) NULL, LSEDatabase sysname NULL,
    LSEOrigem nvarchar(400) NULL, LSEMensagem nvarchar(2000) NULL);

IF @K_LSPrim + @K_LSSec + @K_LSMon = 0
BEGIN
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    VALUES (N'23 - Log Shipping', N'Log shipping', N'NÃO CONFIGURADO', N'INFO',
            N'Nenhuma database primária, secundária ou monitorada em msdb.dbo.log_shipping_*.');
END
ELSE
BEGIN
    INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
    VALUES (N'23 - Log Shipping', N'Log shipping', N'CONFIGURADO', N'INFO',
            CONVERT(nvarchar(10), @K_LSPrim) + N' database(s) primária(s) e ' + CONVERT(nvarchar(10), @K_LSSec)
            + N' secundária(s) nesta instância. Detalhes na guia Alta Disponibilidade.');

    /* ---------- jobs do log shipping ---------- */
    BEGIN TRY
        INSERT INTO #V11_LSJob (job_id, JobName, JobEnabled, JobLastRun, JobLastOutcome, JobNextRun)
        SELECT j.job_id, j.name, CONVERT(bit, j.enabled),
               CASE WHEN lh.run_date > 0 THEN msdb.dbo.agent_datetime(lh.run_date, lh.run_time) END,
               CASE lh.run_status WHEN 0 THEN N'Falha' WHEN 1 THEN N'Sucesso' WHEN 2 THEN N'Nova tentativa'
                                  WHEN 3 THEN N'Cancelado' WHEN 4 THEN N'Em execução' ELSE N'Sem histórico' END,
               nx.NextRun
        FROM msdb.dbo.sysjobs j
        OUTER APPLY (SELECT TOP (1) h.run_status, h.run_date, h.run_time
                     FROM msdb.dbo.sysjobhistory h
                     WHERE h.job_id = j.job_id AND h.step_id = 0
                     ORDER BY h.instance_id DESC) lh
        OUTER APPLY (SELECT MIN(msdb.dbo.agent_datetime(js.next_run_date, js.next_run_time)) AS NextRun
                     FROM msdb.dbo.sysjobschedules js
                     WHERE js.job_id = j.job_id AND js.next_run_date > 0) nx
        WHERE j.job_id IN (SELECT backup_job_id FROM msdb.dbo.log_shipping_primary_databases
                           UNION SELECT copy_job_id FROM msdb.dbo.log_shipping_secondary
                           UNION SELECT restore_job_id FROM msdb.dbo.log_shipping_secondary)
           OR j.name LIKE N'LSAlert[_]%' OR j.name LIKE N'LSBackup[_]%'
           OR j.name LIKE N'LSCopy[_]%'  OR j.name LIKE N'LSRestore[_]%';

        INSERT INTO #D25_LSJobs (LSJFuncao, LSJJob, LSJDatabase, LSJHabilitado, LSJUltimaExecucao, LSJResultado, LSJProximaExecucao, LSJStatus)
        SELECT
            CASE WHEN EXISTS (SELECT 1 FROM msdb.dbo.log_shipping_primary_databases x WHERE x.backup_job_id = v.job_id) THEN N'Backup'
                 WHEN EXISTS (SELECT 1 FROM msdb.dbo.log_shipping_secondary x WHERE x.copy_job_id = v.job_id) THEN N'Cópia'
                 WHEN EXISTS (SELECT 1 FROM msdb.dbo.log_shipping_secondary x WHERE x.restore_job_id = v.job_id) THEN N'Restore'
                 WHEN v.JobName LIKE N'LSAlert[_]%' THEN N'Alerta'
                 ELSE N'Outro' END,
            v.JobName,
            COALESCE((SELECT TOP (1) x.primary_database FROM msdb.dbo.log_shipping_primary_databases x WHERE x.backup_job_id = v.job_id),
                     (SELECT TOP (1) sd.secondary_database FROM msdb.dbo.log_shipping_secondary x
                      JOIN msdb.dbo.log_shipping_secondary_databases sd ON sd.secondary_id = x.secondary_id
                      WHERE v.job_id IN (x.copy_job_id, x.restore_job_id)),
                     N''),
            v.JobEnabled, v.JobLastRun, v.JobLastOutcome, v.JobNextRun,
            CASE WHEN v.JobEnabled = 0 OR v.JobLastOutcome = N'Falha' THEN N'ATENÇÃO' ELSE N'OK' END
        FROM #V11_LSJob v;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'25 - Log shipping (jobs)', ERROR_MESSAGE());
    END CATCH;

    /* ---------- databases primárias ---------- */
    BEGIN TRY
        INSERT INTO #D25_LSPrimario (LSPDatabase, LSPSecundarios, LSPBackupDir, LSPBackupShare, LSPRetencaoMin, LSPCompressao,
            LSPMonitor, LSPJobBackup, LSPUltimoArquivo, LSPUltimoBackup, LSPMinDesdeBackup, LSPLimiteMin, LSPAlertaHabilitado, LSPStatus)
        SELECT pd.primary_database,
               LEFT(sec.Lista, 2000),
               pd.backup_directory,
               pd.backup_share,
               pd.backup_retention_period,
               CASE pd.backup_compression WHEN 0 THEN N'Desligada' WHEN 1 THEN N'Ligada' ELSE N'Padrão da instância' END,
               ISNULL(NULLIF(pd.monitor_server, N''), N'(sem monitor)'),
               bj.JobName,
               COALESCE(mp.last_backup_file, pd.last_backup_file),
               ub.UltimoBackup,
               DATEDIFF(minute, ub.UltimoBackup, GETDATE()),
               ISNULL(mp.backup_threshold, 60),
               mp.threshold_alert_enabled,
               CASE WHEN ub.UltimoBackup IS NULL THEN N'CRÍTICO'
                    WHEN DATEDIFF(minute, ub.UltimoBackup, GETDATE()) > ISNULL(mp.backup_threshold, 60) THEN N'CRÍTICO'
                    WHEN bj.JobEnabled = 0 OR bj.JobLastOutcome = N'Falha' OR bj.job_id IS NULL THEN N'ATENÇÃO'
                    ELSE N'OK' END
        FROM msdb.dbo.log_shipping_primary_databases pd
        LEFT JOIN msdb.dbo.log_shipping_monitor_primary mp ON mp.primary_id = pd.primary_id
        LEFT JOIN #V11_LSJob bj ON bj.job_id = pd.backup_job_id
        OUTER APPLY (SELECT CASE WHEN pd.last_backup_date IS NULL OR mp.last_backup_date > pd.last_backup_date
                                 THEN mp.last_backup_date ELSE pd.last_backup_date END AS UltimoBackup) ub
        OUTER APPLY (SELECT STRING_AGG(CONVERT(nvarchar(max), ps.secondary_server + N'.' + ps.secondary_database), N', ') AS Lista
                     FROM msdb.dbo.log_shipping_primary_secondaries ps
                     WHERE ps.primary_id = pd.primary_id) sec;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'25 - Log shipping (primário)', ERROR_MESSAGE());
    END CATCH;

    /* ---------- databases secundárias ---------- */
    BEGIN TRY
        INSERT INTO #D25_LSSecundario (LSSDatabase, LSSServidorPrimario, LSSDatabasePrimaria, LSSPastaOrigem, LSSPastaDestino,
            LSSRetencaoMin, LSSModoRestore, LSSDesconectaUsuarios, LSSAtrasoRestoreMin, LSSMonitor, LSSJobCopia, LSSJobRestore,
            LSSUltimaCopia, LSSUltimoRestore, LSSUltimoArquivo, LSSMinDesdeRestore, LSSLatenciaMin, LSSLimiteMin,
            LSSAlertaHabilitado, LSSStatus)
        SELECT sd.secondary_database,
               s.primary_server,
               s.primary_database,
               s.backup_source_directory,
               s.backup_destination_directory,
               s.file_retention_period,
               CASE sd.restore_mode WHEN 1 THEN N'STANDBY (leitura)' ELSE N'NORECOVERY' END,
               sd.disconnect_users,
               sd.restore_delay,
               ISNULL(NULLIF(s.monitor_server, N''), N'(sem monitor)'),
               cj.JobName,
               rj.JobName,
               CASE WHEN s.last_copied_date IS NULL OR ms.last_copied_date > s.last_copied_date
                    THEN ms.last_copied_date ELSE s.last_copied_date END,
               ur.UltimoRestore,
               COALESCE(ms.last_restored_file, sd.last_restored_file),
               DATEDIFF(minute, ur.UltimoRestore, GETDATE()),
               ms.last_restored_latency,
               ISNULL(ms.restore_threshold, 45),
               ms.threshold_alert_enabled,
               CASE WHEN ur.UltimoRestore IS NULL THEN N'CRÍTICO'
                    WHEN DATEDIFF(minute, ur.UltimoRestore, GETDATE()) > ISNULL(ms.restore_threshold, 45) + ISNULL(sd.restore_delay, 0) THEN N'CRÍTICO'
                    WHEN ms.last_restored_latency > ISNULL(ms.restore_threshold, 45) + ISNULL(sd.restore_delay, 0) THEN N'CRÍTICO'
                    WHEN cj.JobEnabled = 0 OR rj.JobEnabled = 0 OR cj.JobLastOutcome = N'Falha' OR rj.JobLastOutcome = N'Falha'
                         OR cj.job_id IS NULL OR rj.job_id IS NULL THEN N'ATENÇÃO'
                    ELSE N'OK' END
        FROM msdb.dbo.log_shipping_secondary_databases sd
        JOIN msdb.dbo.log_shipping_secondary s ON s.secondary_id = sd.secondary_id
        LEFT JOIN msdb.dbo.log_shipping_monitor_secondary ms
               ON ms.secondary_id = sd.secondary_id AND ms.secondary_database = sd.secondary_database
        LEFT JOIN #V11_LSJob cj ON cj.job_id = s.copy_job_id
        LEFT JOIN #V11_LSJob rj ON rj.job_id = s.restore_job_id
        OUTER APPLY (SELECT CASE WHEN sd.last_restored_date IS NULL OR ms.last_restored_date > sd.last_restored_date
                                 THEN ms.last_restored_date ELSE sd.last_restored_date END AS UltimoRestore) ur;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'25 - Log shipping (secundário)', ERROR_MESSAGE());
    END CATCH;

    /* ---------- servidor monitor: databases de outras instâncias ---------- */
    BEGIN TRY
        INSERT INTO #D25_LSMonitor (LSMPapel, LSMServidor, LSMDatabase, LSMParceiro, LSMUltimaOperacao, LSMMinDesde,
            LSMLatenciaMin, LSMLimiteMin, LSMAlertaHabilitado, LSMStatus)
        SELECT N'Primário', mp.primary_server, mp.primary_database, N'',
               mp.last_backup_date,
               DATEDIFF(minute, mp.last_backup_date_utc, GETUTCDATE()),
               NULL,
               mp.backup_threshold,
               mp.threshold_alert_enabled,
               CASE WHEN mp.last_backup_date_utc IS NULL
                      OR DATEDIFF(minute, mp.last_backup_date_utc, GETUTCDATE()) > mp.backup_threshold THEN N'CRÍTICO'
                    ELSE N'OK' END
        FROM msdb.dbo.log_shipping_monitor_primary mp
        WHERE NOT EXISTS (SELECT 1 FROM msdb.dbo.log_shipping_primary_databases pd WHERE pd.primary_id = mp.primary_id)
        UNION ALL
        SELECT N'Secundário', ms.secondary_server, ms.secondary_database,
               ms.primary_server + N'.' + ms.primary_database,
               ms.last_restored_date,
               DATEDIFF(minute, ms.last_restored_date_utc, GETUTCDATE()),
               ms.last_restored_latency,
               ms.restore_threshold,
               ms.threshold_alert_enabled,
               CASE WHEN ms.last_restored_date_utc IS NULL
                      OR DATEDIFF(minute, ms.last_restored_date_utc, GETUTCDATE()) > ms.restore_threshold
                      OR ms.last_restored_latency > ms.restore_threshold THEN N'CRÍTICO'
                    ELSE N'OK' END
        FROM msdb.dbo.log_shipping_monitor_secondary ms
        WHERE NOT EXISTS (SELECT 1 FROM msdb.dbo.log_shipping_secondary s WHERE s.secondary_id = ms.secondary_id);
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'25 - Log shipping (monitor)', ERROR_MESSAGE());
    END CATCH;

    /* ---------- erros dos últimos 7 dias ---------- */
    BEGIN TRY
        INSERT INTO #D25_LSErros (LSEData, LSEAgente, LSEDatabase, LSEOrigem, LSEMensagem)
        SELECT TOP (50)
               ed.log_time,
               CASE ed.agent_type WHEN 0 THEN N'Backup' WHEN 1 THEN N'Cópia' WHEN 2 THEN N'Restore' ELSE N'-' END,
               ed.database_name,
               CONVERT(nvarchar(400), ed.source),
               CONVERT(nvarchar(2000), REPLACE(REPLACE(ed.message, NCHAR(13), N' '), NCHAR(10), N' '))
        FROM msdb.dbo.log_shipping_monitor_error_detail ed
        WHERE ed.log_time_utc >= DATEADD(day, -7, GETUTCDATE())
        ORDER BY ed.log_time_utc DESC;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'25 - Log shipping (erros)', ERROR_MESSAGE());
    END CATCH;

    /* ---------- verificações (checklist) ---------- */
    BEGIN TRY
        IF @K_LSPrim > 0
            INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
            SELECT N'23 - Log Shipping', N'Primárias com backup de log acima do limite', CONVERT(nvarchar(30), COUNT(*)),
                   CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
                   N'Minutos desde o último backup de log maior que o limite do monitor (backup_threshold, padrão 60 min) ou nenhum backup registrado.'
            FROM #D25_LSPrimario WHERE LSPStatus = N'CRÍTICO';

        IF @K_LSSec > 0
            INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
            SELECT N'23 - Log Shipping', N'Secundárias com restore acima do limite', CONVERT(nvarchar(30), COUNT(*)),
                   CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
                   N'Tempo desde o último restore ou latência maior que restore_threshold (padrão 45 min) + restore_delay. O secundário está atrasado em relação ao primário (RPO em risco).'
            FROM #D25_LSSecundario WHERE LSSStatus = N'CRÍTICO';

        IF EXISTS (SELECT 1 FROM #D25_LSMonitor)
            INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
            SELECT N'23 - Log Shipping', N'Databases monitoradas (outras instâncias) fora do limite', CONVERT(nvarchar(30), COUNT(*)),
                   CASE WHEN COUNT(*) > 0 THEN N'CRÍTICO' ELSE N'OK' END,
                   N'Esta instância é o servidor monitor de log shipping de outras instâncias.'
            FROM #D25_LSMonitor WHERE LSMStatus = N'CRÍTICO';

        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT N'23 - Log Shipping', N'Erros de log shipping nas últimas 24 h', CONVERT(nvarchar(30), COUNT(*)),
               CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
               N'msdb.dbo.log_shipping_monitor_error_detail. As mensagens dos últimos 7 dias estão no detalhamento.'
        FROM msdb.dbo.log_shipping_monitor_error_detail
        WHERE log_time_utc >= DATEADD(hour, -24, GETUTCDATE());

        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT N'23 - Log Shipping', N'Jobs de log shipping desabilitados', CONVERT(nvarchar(30), COUNT(*)),
               CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
               N'Jobs LSBackup, LSCopy, LSRestore e LSAlert. Job desabilitado interrompe o envio, a cópia ou o restore dos logs.'
        FROM #D25_LSJobs WHERE LSJHabilitado = 0;

        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT N'23 - Log Shipping', N'Jobs de log shipping com falha na última execução', CONVERT(nvarchar(30), COUNT(*)),
               CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
               N'Ver o histórico do job e os erros de log shipping no detalhamento.'
        FROM #D25_LSJobs WHERE LSJResultado = N'Falha';

        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT N'23 - Log Shipping', N'Configurações sem servidor monitor', CONVERT(nvarchar(30), COUNT(*)),
               CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
               N'Sem monitor não há job LSAlert nem os erros 14420/14421 de atraso: o atraso só é percebido manualmente.'
        FROM (SELECT monitor_server FROM msdb.dbo.log_shipping_primary_databases
              UNION ALL
              SELECT monitor_server FROM msdb.dbo.log_shipping_secondary) x
        WHERE ISNULL(x.monitor_server, N'') = N'';

        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT N'23 - Log Shipping', N'Alertas de limite (threshold) desabilitados', CONVERT(nvarchar(30), COUNT(*)),
               CASE WHEN COUNT(*) > 0 THEN N'ATENÇÃO' ELSE N'OK' END,
               N'threshold_alert_enabled = 0 no monitor: o atraso acima do limite não gera os erros 14420/14421.'
        FROM (SELECT threshold_alert_enabled AS a FROM msdb.dbo.log_shipping_monitor_primary
              UNION ALL
              SELECT threshold_alert_enabled FROM msdb.dbo.log_shipping_monitor_secondary) x
        WHERE x.a = 0;

        INSERT INTO #AssessmentSummary (Categoria, Indicador, Valor, Status, Observacao)
        SELECT N'23 - Log Shipping', N'Alertas do SQL Agent para os erros 14420/14421', CONVERT(nvarchar(30), COUNT(*)),
               CASE WHEN COUNT(*) = 0 THEN N'ATENÇÃO' ELSE N'OK' END,
               N'O job LSAlert grava os erros 14420 (backup atrasado) e 14421 (restore atrasado) no errorlog; sem alerta do Agent associado a um operador, ninguém é avisado.'
        FROM msdb.dbo.sysalerts
        WHERE message_id IN (14420, 14421) AND enabled = 1;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'25 - Log shipping (verificações)', ERROR_MESSAGE());
    END CATCH;
END;

/* ============================================================================
   RESULTADO CENTRALIZADO
============================================================================ */
PRINT '=================================================================';
PRINT 'SQL SERVER - ASSESSMENT / DISCOVERY';
PRINT 'Servidor: ' + ISNULL(@ServerName,'N/D');
PRINT 'Data: ' + CONVERT(varchar(19),@AssessmentDate,120);
PRINT '=================================================================';

IF @SomenteHTML = 0
    SELECT
        ID,
        Categoria,
        Indicador,
        Valor,
        Status,
        Observacao
    FROM #AssessmentSummary
    ORDER BY ID;

/* ============================================================================
   COLETA DOS DETALHAMENTOS
   Cada seção é gravada em uma tabela temporária e registrada em #Secoes.
   A exibição (SSMS) e o HTML são gerados no final, na mesma ordem.
============================================================================ */

/* DETALHAMENTO 01 - INSTÂNCIA */
BEGIN TRY
    SELECT
        SERVERPROPERTY('MachineName') AS MachineName,
        SERVERPROPERTY('ServerName') AS ServerName,
        SERVERPROPERTY('InstanceName') AS InstanceName,
        SERVERPROPERTY('IsClustered') AS IsClustered,
        SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS ComputerNamePhysicalNetBIOS,
        SERVERPROPERTY('Edition') AS Edition,
        SERVERPROPERTY('ProductLevel') AS ProductLevel,
        SERVERPROPERTY('ProductUpdateLevel') AS ProductUpdateLevel,
        SERVERPROPERTY('ProductVersion') AS ProductVersion,
        SERVERPROPERTY('ProductBuild') AS ProductBuild,
        SERVERPROPERTY('ProductBuildType') AS ProductBuildType,
        SERVERPROPERTY('ProductUpdateReference') AS ProductUpdateReference,
        SERVERPROPERTY('ProcessID') AS ProcessID,
        SERVERPROPERTY('Collation') AS Collation,
        SERVERPROPERTY('IsFullTextInstalled') AS IsFullTextInstalled,
        SERVERPROPERTY('IsIntegratedSecurityOnly') AS IsIntegratedSecurityOnly,
        SERVERPROPERTY('FilestreamConfiguredLevel') AS FilestreamConfiguredLevel,
        SERVERPROPERTY('IsHadrEnabled') AS IsHadrEnabled,
        SERVERPROPERTY('HadrManagerStatus') AS HadrManagerStatus,
        SERVERPROPERTY('InstanceDefaultDataPath') AS InstanceDefaultDataPath,
        SERVERPROPERTY('InstanceDefaultLogPath') AS InstanceDefaultLogPath,
        SERVERPROPERTY('ErrorLogFileName') AS ErrorLogFileName,
        SERVERPROPERTY('BuildClrVersion') AS BuildClrVersion,
        SERVERPROPERTY('IsXTPSupported') AS IsXTPSupported,
        SERVERPROPERTY('IsPolybaseInstalled') AS IsPolybaseInstalled
    INTO #D01_Instancia;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'01', N'Propriedades da instância', N'#D01_Instancia', N'ServerName', N'SERVERPROPERTY: versão, edição, collation, caminhos padrão e recursos instalados.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'01 - Propriedades da instância', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 02 - HARDWARE / SO / SERVIÇOS */
BEGIN TRY
    SELECT cpu_count AS LogicalCPUCount,
           scheduler_count,
           socket_count,
           cores_per_socket,
           numa_node_count,
           physical_memory_kb/1024 AS PhysicalMemoryMB,
           max_workers_count,
           affinity_type_desc,
           sqlserver_start_time,
           DATEDIFF(hour,sqlserver_start_time,GETDATE()) AS UptimeHours,
           virtual_machine_type_desc,
           softnuma_configuration_desc,
           sql_memory_model_desc
    INTO #D02_Hardware
    FROM sys.dm_os_sys_info;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'02', N'Hardware (CPU, memória, NUMA)', N'#D02_Hardware', N'LogicalCPUCount', N'sys.dm_os_sys_info.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'02 - Hardware', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT host_platform, host_distribution, host_release, host_service_pack_level,
           host_sku, os_language_version
    INTO #D02_HostInfo
    FROM sys.dm_os_host_info;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'02', N'Sistema operacional', N'#D02_HostInfo', N'host_platform', N'sys.dm_os_host_info.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'02 - Sistema operacional', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT servicename, process_id, startup_type_desc, status_desc,
           last_startup_time, service_account, is_clustered, cluster_nodename,
           filename, instant_file_initialization_enabled
    INTO #D02_Servicos
    FROM sys.dm_server_services;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'02', N'Serviços do SQL Server', N'#D02_Servicos', N'servicename', N'Contas de serviço, tipo de inicialização e Instant File Initialization.', N'filename');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'02 - Serviços', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 03 - CLUSTER / FCI */
IF @IsClustered = 1
BEGIN
    BEGIN TRY
        SELECT * INTO #D03_ClusterNodes FROM sys.dm_os_cluster_nodes;
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'03', N'Nós do cluster (FCI)', N'#D03_ClusterNodes', N'NodeName', N'sys.dm_os_cluster_nodes.');

        SELECT * INTO #D03_ClusterProps FROM sys.dm_os_cluster_properties;
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'03', N'Propriedades do cluster (FCI)', N'#D03_ClusterProps', N'VerboseLogging', N'sys.dm_os_cluster_properties.');
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'03 - Cluster / FCI', ERROR_MESSAGE());
    END CATCH;
END
ELSE
BEGIN
    INSERT INTO #Secoes (Grupo, Titulo, Nota)
    VALUES (N'03', N'Cluster / FCI', N'FCI não configurado para esta instância.');
END;

/* DETALHAMENTO 04 - ALWAYS ON */
IF @IsHadrEnabled = 1
BEGIN
    BEGIN TRY
        SELECT ag.name AS AGName,
               ag.group_id,
               ag.failure_condition_level,
               ag.health_check_timeout,
               ag.db_failover,
               ag.dtc_support,
               ag.basic_features,
               ag.cluster_type_desc,
               ag.required_synchronized_secondaries_to_commit
        INTO #D04_AGs
        FROM sys.availability_groups ag;
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'04', N'Availability Groups', N'#D04_AGs', N'AGName', N'Configuração de cada AG.');

        SELECT ag.name AS AGName,
               ar.replica_server_name,
               ar.endpoint_url,
               ar.availability_mode_desc,
               ar.failover_mode_desc,
               ar.seeding_mode_desc,
               ar.session_timeout,
               ar.primary_role_allow_connections_desc,
               ar.secondary_role_allow_connections_desc,
               ar.backup_priority,
               ars.role_desc,
               ars.operational_state_desc,
               ars.connected_state_desc,
               ars.recovery_health_desc,
               ars.synchronization_health_desc
        INTO #D04_Replicas
        FROM sys.availability_replicas ar
        JOIN sys.availability_groups ag ON ag.group_id=ar.group_id
        LEFT JOIN sys.dm_hadr_availability_replica_states ars
          ON ars.group_id=ar.group_id AND ars.replica_id=ar.replica_id;
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'04', N'Réplicas dos AGs', N'#D04_Replicas', N'AGName, replica_server_name', N'Modo de disponibilidade, failover, seeding e saúde de cada réplica.');

        SELECT ag.name AS AGName,
               ar.replica_server_name,
               adc.database_name,
               drs.is_local,
               drs.is_primary_replica,
               drs.synchronization_state_desc,
               drs.synchronization_health_desc,
               drs.database_state_desc,
               drs.log_send_queue_size,
               drs.log_send_rate,
               drs.redo_queue_size,
               drs.redo_rate,
               drs.last_sent_time,
               drs.last_received_time,
               drs.last_hardened_time,
               drs.last_redone_time,
               drs.last_commit_time
        INTO #D04_Databases
        FROM sys.dm_hadr_database_replica_states drs
        JOIN sys.availability_databases_cluster adc
          ON adc.group_id=drs.group_id AND adc.group_database_id=drs.group_database_id
        JOIN sys.availability_groups ag ON ag.group_id=drs.group_id
        JOIN sys.availability_replicas ar
          ON ar.group_id=drs.group_id AND ar.replica_id=drs.replica_id;
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'04', N'Databases nos AGs', N'#D04_Databases', N'AGName, replica_server_name, database_name', N'Estado de sincronização, filas de envio/redo (KB) e últimos tempos por database/réplica.');

        SELECT cluster_name, quorum_type_desc, quorum_state_desc
        INTO #D04_Cluster
        FROM sys.dm_hadr_cluster;
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'04', N'Cluster WSFC / quórum', N'#D04_Cluster', N'cluster_name', N'sys.dm_hadr_cluster.');
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'04 - Always On', ERROR_MESSAGE());
    END CATCH;
END
ELSE
BEGIN
    INSERT INTO #Secoes (Grupo, Titulo, Nota)
    VALUES (N'04', N'Always On', N'Always On não habilitado nesta instância.');
END;

/* DETALHAMENTO 05 - LINKED SERVERS */
BEGIN TRY
    SELECT name,
           product,
           provider,
           data_source,
           location,
           catalog,
           connect_timeout,
           query_timeout,
           is_linked,
           is_data_access_enabled,
           is_rpc_out_enabled,
           is_remote_login_enabled,
           is_system
    INTO #D05_LinkedServers
    FROM sys.servers
    WHERE is_linked = 1;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'05', N'Linked Servers', N'#D05_LinkedServers', N'name', N'Provider, data source e opções de RPC/data access.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'05 - Linked Servers', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 06 - ENDPOINTS
   sys.endpoints não expõe diretamente porta/IP.
   Porta/IP TCP: sys.tcp_endpoints | role_desc: sys.database_mirroring_endpoints */
BEGIN TRY
    SELECT e.name,
           e.endpoint_id,
           e.type_desc,
           e.state_desc,
           e.protocol_desc,
           dme.role_desc AS database_mirroring_role_desc,
           dme.is_encryption_enabled,
           dme.connection_auth_desc,
           tcp.port,
           tcp.is_dynamic_port,
           tcp.ip_address
    INTO #D06_Endpoints
    FROM sys.endpoints AS e
    LEFT JOIN sys.tcp_endpoints AS tcp
           ON tcp.endpoint_id = e.endpoint_id
    LEFT JOIN sys.database_mirroring_endpoints AS dme
           ON dme.endpoint_id = e.endpoint_id;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'06', N'Endpoints', N'#D06_Endpoints', N'endpoint_id', N'Endpoints TCP, HADR e Service Broker.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'06 - Endpoints', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 07 - DATABASES */
BEGIN TRY
    SELECT d.database_id,
           d.name AS DatabaseName,
           d.state_desc,
           d.user_access_desc,
           SUSER_SNAME(d.owner_sid) AS DatabaseOwner,
           d.recovery_model_desc,
           d.compatibility_level,
           d.page_verify_option_desc,
           d.is_auto_close_on,
           d.is_auto_shrink_on,
           d.is_auto_create_stats_on,
           d.is_auto_update_stats_on,
           d.is_auto_update_stats_async_on,
           d.is_parameterization_forced,
           d.snapshot_isolation_state_desc,
           d.is_read_committed_snapshot_on,
           d.is_cdc_enabled,
           d.is_encrypted,
           d.is_query_store_on,
           d.log_reuse_wait_desc,
           CONVERT(decimal(18,2), SUM(CASE WHEN mf.type=0 THEN CONVERT(bigint, mf.size) ELSE 0 END)/128.0) AS DataSizeMB,
           CONVERT(decimal(18,2), SUM(CASE WHEN mf.type=1 THEN CONVERT(bigint, mf.size) ELSE 0 END)/128.0) AS LogSizeMB,
           COUNT(CASE WHEN mf.type=0 THEN 1 END) AS DataFiles,
           COUNT(CASE WHEN mf.type=1 THEN 1 END) AS LogFiles
    INTO #D07_Databases
    FROM sys.databases d
    LEFT JOIN sys.master_files mf ON mf.database_id=d.database_id
    GROUP BY d.database_id,d.name,d.state_desc,d.user_access_desc,d.owner_sid,
             d.recovery_model_desc,d.compatibility_level,d.page_verify_option_desc,
             d.is_auto_close_on,d.is_auto_shrink_on,d.is_auto_create_stats_on,
             d.is_auto_update_stats_on,d.is_auto_update_stats_async_on,
             d.is_parameterization_forced,d.snapshot_isolation_state_desc,
             d.is_read_committed_snapshot_on,d.is_cdc_enabled,d.is_encrypted,
             d.is_query_store_on,d.log_reuse_wait_desc;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'07', N'Databases', N'#D07_Databases', N'database_id', N'Estado, recovery model, compatibilidade, opções e tamanho de cada database.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'07 - Databases', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 08 - ARQUIVOS / VOLUMES */
BEGIN TRY
    SELECT DB_NAME(database_id) AS DatabaseName,
           file_id,
           name AS LogicalName,
           physical_name,
           type_desc,
           state_desc,
           CONVERT(decimal(18,2), size/128.0) AS SizeMB,
           is_percent_growth,
           CONVERT(decimal(18,2), CASE WHEN is_percent_growth=1 THEN growth ELSE growth/128.0 END) AS GrowthValue,
           CASE WHEN is_percent_growth=1 THEN '%' ELSE 'MB' END AS GrowthUnit,
           max_size
    INTO #D08_Arquivos
    FROM sys.master_files;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'08', N'Arquivos de dados e log', N'#D08_Arquivos', N'DatabaseName, file_id', N'Localização, tamanho e configuração de crescimento de cada arquivo.', N'physical_name');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'08 - Arquivos', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT DISTINCT
           vs.volume_mount_point,
           vs.file_system_type,
           vs.logical_volume_name,
           CONVERT(decimal(18,2), vs.total_bytes/1073741824.0) AS TotalGB,
           CONVERT(decimal(18,2), vs.available_bytes/1073741824.0) AS AvailableGB,
           CONVERT(decimal(9,2), vs.available_bytes*100.0/NULLIF(vs.total_bytes,0)) AS FreePct,
           vs.supports_compression,
           vs.is_compressed,
           vs.supports_sparse_files,
           vs.supports_alternate_streams
    INTO #D08_Volumes
    FROM sys.master_files mf
    CROSS APPLY sys.dm_os_volume_stats(mf.database_id,mf.file_id) vs;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'08', N'Volumes com arquivos do SQL Server', N'#D08_Volumes', N'volume_mount_point', N'sys.dm_os_volume_stats.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'08 - Volumes', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT fixed_drive_path,
           drive_type_desc,
           CONVERT(decimal(18,2), free_space_in_bytes/1073741824.0) AS AvailableGB
    INTO #D08_Drives
    FROM sys.dm_os_enumerate_fixed_drives;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'08', N'Unidades fixas do servidor', N'#D08_Drives', N'fixed_drive_path', N'sys.dm_os_enumerate_fixed_drives (Windows).');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'08 - Unidades fixas', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 09 - TEMPDB */
BEGIN TRY
    SELECT name,
           physical_name,
           type_desc,
           CONVERT(decimal(18,2), size/128.0) AS SizeMB,
           is_percent_growth,
           CONVERT(decimal(18,2), CASE WHEN is_percent_growth=1 THEN growth ELSE growth/128.0 END) AS GrowthValue,
           CASE WHEN is_percent_growth=1 THEN '%' ELSE 'MB' END AS GrowthUnit,
           file_id
    INTO #D09_TempdbArquivos
    FROM tempdb.sys.database_files;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'09', N'TempDB - arquivos', N'#D09_TempdbArquivos', N'type_desc DESC, file_id', N'Quantidade, tamanho e crescimento dos arquivos do TempDB.', N'physical_name');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'09 - TempDB', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT DB_NAME(database_id) AS DatabaseName,
           reserved_page_count,
           CONVERT(decimal(18,2), reserved_space_kb/1024.0) AS ReservedSpaceMB
    INTO #D09_VersionStore
    FROM sys.dm_tran_version_store_space_usage;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'09', N'TempDB - version store por database', N'#D09_VersionStore', N'ReservedSpaceMB DESC', N'Espaço de version store (RCSI/SI, AG legível, triggers, online index).');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'09 - Version store', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 10 - BACKUPS */
BEGIN TRY
    SELECT d.name AS DatabaseName,
           d.recovery_model_desc,
           d.state_desc,
           MAX(CASE WHEN b.type='D' THEN b.backup_finish_date END) AS LastFullBackup,
           MAX(CASE WHEN b.type='I' THEN b.backup_finish_date END) AS LastDifferentialBackup,
           MAX(CASE WHEN b.type='L' THEN b.backup_finish_date END) AS LastLogBackup,
           CONVERT(datetime, DATABASEPROPERTYEX(d.name,'LastGoodCheckDbTime')) AS LastGoodCheckDB
    INTO #D10_Backups
    FROM sys.databases d
    LEFT JOIN msdb.dbo.backupset b
      ON b.database_name=d.name
     AND b.backup_finish_date >= DATEADD(day,-30,GETDATE())
    GROUP BY d.name,d.recovery_model_desc,d.state_desc;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'10', N'Backups (últimos 30 dias) e último CHECKDB', N'#D10_Backups', N'DatabaseName', N'Datas vazias = sem backup nos últimos 30 dias. LastGoodCheckDB 1900-01-01 = CHECKDB nunca executado com sucesso.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'10 - Backups', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 11 - SQL AGENT / ALERTAS */
BEGIN TRY
    SELECT j.name AS JobName,
           j.enabled AS JobEnabled,
           SUSER_SNAME(j.owner_sid) AS JobOwner,
           c.name AS CategoryName,
           j.date_created,
           h.run_status AS LastRunStatus,
           CASE h.run_status WHEN 0 THEN N'Falha' WHEN 1 THEN N'Sucesso' WHEN 2 THEN N'Nova tentativa'
                             WHEN 3 THEN N'Cancelado' WHEN 4 THEN N'Em execução' END AS LastRunStatusDesc,
           msdb.dbo.agent_datetime(h.run_date,h.run_time) AS LastRunDateTime,
           h.run_duration AS LastRunDuration,
           h.message AS LastRunMessage
    INTO #D11_Jobs
    FROM msdb.dbo.sysjobs j
    LEFT JOIN msdb.dbo.syscategories c ON c.category_id=j.category_id
    OUTER APPLY
    (
        SELECT TOP (1) h1.run_status,h1.run_date,h1.run_time,h1.run_duration,h1.message
        FROM msdb.dbo.sysjobhistory h1
        WHERE h1.job_id=j.job_id AND h1.step_id=0
        ORDER BY h1.instance_id DESC
    ) h;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'11', N'Jobs do SQL Agent', N'#D11_Jobs', N'JobEnabled DESC, JobName', N'Owner, categoria e resultado da última execução. LastRunDuration no formato HHMMSS.', N'LastRunMessage');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'11 - Jobs', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT name,event_source,message_id,severity,enabled,has_notification,
           delay_between_responses,occurrence_count,last_occurrence_date,last_occurrence_time
    INTO #D11_Alertas
    FROM msdb.dbo.sysalerts;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'11', N'Alertas do SQL Agent', N'#D11_Alertas', N'name', N'msdb.dbo.sysalerts.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'11 - Alertas', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 12 - SEGURANÇA */
BEGIN TRY
    SELECT sp.name AS LoginName,
           sp.type_desc,
           sp.is_disabled,
           sp.default_database_name,
           sp.create_date,
           sp.modify_date,
           CASE WHEN EXISTS
           (
               SELECT 1
               FROM sys.server_role_members srm
               JOIN sys.server_principals rp ON rp.principal_id=srm.role_principal_id
               WHERE srm.member_principal_id=sp.principal_id AND rp.name='sysadmin'
           ) THEN 1 ELSE 0 END AS IsSysadmin
    INTO #D12_Logins
    FROM sys.server_principals sp
    WHERE sp.type IN ('S','U','G');
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'12', N'Logins', N'#D12_Logins', N'IsSysadmin DESC, LoginName', N'Logins SQL, usuários e grupos Windows. IsSysadmin = 1 indica membro da role sysadmin.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'12 - Logins', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 13 - CONFIGURAÇÕES */
BEGIN TRY
    SELECT name,
           value AS ConfiguredValue,
           value_in_use AS RunningValue,
           minimum,
           maximum,
           is_dynamic,
           is_advanced,
           description
    INTO #D13_Config
    FROM sys.configurations
    WHERE name IN
    (
        'max server memory (MB)',
        'min server memory (MB)',
        'max degree of parallelism',
        'cost threshold for parallelism',
        'backup compression default',
        'optimize for ad hoc workloads',
        'remote admin connections',
        'remote access',
        'contained database authentication',
        'clr enabled',
        'xp_cmdshell',
        'Ole Automation Procedures',
        'Database Mail XPs',
        'Ad Hoc Distributed Queries',
        'default trace enabled'
    );
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'13', N'Configurações relevantes (sp_configure)', N'#D13_Config', N'name', N'Principais configurações de desempenho e superfície de ataque.', N'description');

    SELECT configuration_id,name,value_in_use,description
    INTO #D13_ConfigPendente
    FROM sys.configurations
    WHERE value_in_use <> value;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'13', N'Configurações pendentes (RECONFIGURE/restart)', N'#D13_ConfigPendente', N'name', N'Configurações cujo valor configurado difere do valor em uso.', N'description');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'13 - Configurações', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 14 - MEMÓRIA / CPU / PLE */
BEGIN TRY
    SELECT CONVERT(decimal(18,2), physical_memory_in_use_kb/1024.0) AS SQLMemoryUsedMB,
           CONVERT(decimal(18,2), locked_page_allocations_kb/1024.0) AS LockedPagesMB,
           CONVERT(decimal(18,2), large_page_allocations_kb/1024.0) AS LargePagesMB,
           page_fault_count,
           memory_utilization_percentage,
           CONVERT(decimal(18,2), available_commit_limit_kb/1024.0) AS AvailableCommitLimitMB,
           process_physical_memory_low,
           process_virtual_memory_low
    INTO #D14_ProcessMemory
    FROM sys.dm_os_process_memory;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'14', N'Memória do processo SQL Server', N'#D14_ProcessMemory', N'SQLMemoryUsedMB', N'sys.dm_os_process_memory.');

    SELECT CONVERT(decimal(18,2), total_physical_memory_kb/1024.0) AS PhysicalMemoryMB,
           CONVERT(decimal(18,2), available_physical_memory_kb/1024.0) AS AvailablePhysicalMemoryMB,
           CONVERT(decimal(18,2), total_page_file_kb/1024.0) AS PageFileCommitLimitMB,
           CONVERT(decimal(18,2), available_page_file_kb/1024.0) AS AvailablePageFileMB,
           CONVERT(decimal(18,2), system_cache_kb/1024.0) AS SystemCacheMB,
           system_memory_state_desc
    INTO #D14_SysMemory
    FROM sys.dm_os_sys_memory;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'14', N'Memória do sistema operacional', N'#D14_SysMemory', N'PhysicalMemoryMB', N'sys.dm_os_sys_memory.');

    SELECT RTRIM(object_name) AS ObjectName,
           instance_name,
           cntr_value AS PageLifeExpectancy,
           GETDATE() AS CollectionTime
    INTO #D14_PLE
    FROM sys.dm_os_performance_counters
    WHERE object_name LIKE '%Buffer Node%'
      AND counter_name='Page life expectancy';
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'14', N'Page Life Expectancy por nó NUMA', N'#D14_PLE', N'instance_name', N'Valor pontual; avaliar em conjunto com a tendência.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'14 - Memória', ERROR_MESSAGE());
END CATCH;

BEGIN TRY
    SELECT TOP (256)
           SQLProcessUtilization AS SQLServerProcessCPU,
           SystemIdle AS SystemIdleCPU,
           100-SystemIdle-SQLProcessUtilization AS OtherProcessCPU,
           DATEADD(ms,-1*(ts.ms_ticks-[timestamp]),GETDATE()) AS EventTime
    INTO #D14_CPU
    FROM
    (
        SELECT record.value('(./Record/@id)[1]','int') AS record_id,
               record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int') AS SystemIdle,
               record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS SQLProcessUtilization,
               [timestamp]
        FROM
        (
            SELECT [timestamp],CONVERT(xml,record) AS record
            FROM sys.dm_os_ring_buffers
            WHERE ring_buffer_type=N'RING_BUFFER_SCHEDULER_MONITOR'
              AND record LIKE N'%<SystemHealth>%'
        ) x
    ) y
    CROSS JOIN (SELECT ms_ticks FROM sys.dm_os_sys_info) ts
    ORDER BY record_id DESC;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'14', N'Histórico de CPU (ring buffer, 1 amostra/min)', N'#D14_CPU', N'EventTime DESC', N'Uso de CPU pelo SQL Server e por outros processos nos últimos minutos.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'14 - Histórico de CPU', ERROR_MESSAGE());
END CATCH;

/* DETALHAMENTO 15 - I/O */
BEGIN TRY
    SELECT DB_NAME(fs.database_id) AS DatabaseName,
           mf.file_id,
           mf.physical_name,
           mf.type_desc,
           CONVERT(decimal(18,2), mf.size/128.0) AS FileSizeMB,
           CONVERT(decimal(18,2), CASE WHEN fs.num_of_reads=0 THEN 0 ELSE fs.io_stall_read_ms*1.0/fs.num_of_reads END) AS AvgReadLatencyMS,
           CONVERT(decimal(18,2), CASE WHEN fs.num_of_writes=0 THEN 0 ELSE fs.io_stall_write_ms*1.0/fs.num_of_writes END) AS AvgWriteLatencyMS,
           CONVERT(decimal(18,2), CASE WHEN fs.num_of_reads+fs.num_of_writes=0 THEN 0 ELSE (fs.io_stall_read_ms+fs.io_stall_write_ms)*1.0/(fs.num_of_reads+fs.num_of_writes) END) AS AvgIOLatencyMS,
           fs.num_of_reads,
           fs.num_of_writes,
           fs.io_stall_read_ms,
           fs.io_stall_write_ms
    INTO #D15_IOLatencia
    FROM sys.dm_io_virtual_file_stats(NULL,NULL) fs
    JOIN sys.master_files mf ON mf.database_id=fs.database_id AND mf.file_id=fs.file_id;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'15', N'Latência de I/O por arquivo', N'#D15_IOLatencia', N'AvgIOLatencyMS DESC', N'Médias acumuladas desde o último restart. Referência: dados < 20 ms, log < 5 ms.', N'physical_name');

    ;WITH IOByDB AS
    (
        SELECT database_id,
               SUM(num_of_bytes_read+num_of_bytes_written)/1048576.0 AS TotalIOMB,
               SUM(num_of_bytes_read)/1048576.0 AS ReadIOMB,
               SUM(num_of_bytes_written)/1048576.0 AS WriteIOMB
        FROM sys.dm_io_virtual_file_stats(NULL,NULL)
        GROUP BY database_id
    )
    SELECT ROW_NUMBER() OVER(ORDER BY TotalIOMB DESC) AS IORank,
           DB_NAME(database_id) AS DatabaseName,
           CONVERT(decimal(18,2), TotalIOMB) AS TotalIOMB,
           CONVERT(decimal(18,2), ReadIOMB) AS ReadIOMB,
           CONVERT(decimal(18,2), WriteIOMB) AS WriteIOMB,
           CONVERT(decimal(9,2), TotalIOMB*100.0/NULLIF(SUM(TotalIOMB) OVER(),0)) AS TotalIOPct
    INTO #D15_IOPorDB
    FROM IOByDB;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'15', N'Volume de I/O por database', N'#D15_IOPorDB', N'IORank', N'Distribuição do I/O acumulado desde o último restart.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'15 - I/O', ERROR_MESSAGE());
END CATCH;

/* Avisos de I/O > 15 segundos nos seis logs mais recentes. */
CREATE TABLE #IOWarningResults(LogDate datetime,ProcessInfo nvarchar(128),LogText nvarchar(4000));
BEGIN TRY
    INSERT INTO #IOWarningResults EXEC xp_readerrorlog 0,1,N'taking longer than 15 seconds';
    INSERT INTO #IOWarningResults EXEC xp_readerrorlog 1,1,N'taking longer than 15 seconds';
    INSERT INTO #IOWarningResults EXEC xp_readerrorlog 2,1,N'taking longer than 15 seconds';
    INSERT INTO #IOWarningResults EXEC xp_readerrorlog 3,1,N'taking longer than 15 seconds';
    INSERT INTO #IOWarningResults EXEC xp_readerrorlog 4,1,N'taking longer than 15 seconds';
    INSERT INTO #IOWarningResults EXEC xp_readerrorlog 5,1,N'taking longer than 15 seconds';
END TRY
BEGIN CATCH
    INSERT INTO #IOWarningResults VALUES (GETDATE(),'ASSESSMENT',ERROR_MESSAGE());
END CATCH;
INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
VALUES (N'15', N'Avisos de I/O > 15 s no ErrorLog', N'#IOWarningResults', N'LogDate DESC', N'Ocorrências de "taking longer than 15 seconds" nos seis ErrorLogs mais recentes.', N'LogText');

/* DETALHAMENTO 16 - SUSPECT PAGES / TRACE FLAGS */
BEGIN TRY
    SELECT DB_NAME(sp.database_id) AS DatabaseName,
           sp.file_id,
           sp.page_id,
           sp.event_type,
           sp.error_count,
           sp.last_update_date,
           mf.name AS LogicalName,
           mf.physical_name
    INTO #D16_SuspectPages
    FROM msdb.dbo.suspect_pages sp
    LEFT JOIN sys.master_files mf
      ON mf.database_id=sp.database_id AND mf.file_id=sp.file_id;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'16', N'Suspect pages', N'#D16_SuspectPages', N'last_update_date DESC', N'event_type 1-3 = ativos; 4 = restaurado; 5 = reparado; 7 = desalocado pelo DBCC.', N'physical_name');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'16 - Suspect pages', ERROR_MESSAGE());
END CATCH;

CREATE TABLE #D16_TraceFlags (TraceFlag int, Status int, Global int, Session int);
BEGIN TRY
    INSERT INTO #D16_TraceFlags EXEC ('DBCC TRACESTATUS(-1) WITH NO_INFOMSGS');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'16 - Trace flags', ERROR_MESSAGE());
END CATCH;
INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
VALUES (N'16', N'Trace flags globais ativos', N'#D16_TraceFlags', N'TraceFlag', N'DBCC TRACESTATUS(-1).');

/* DETALHAMENTO 17 - BUFFER POOL POR DATABASE / LOG / VLF */
BEGIN TRY
    ;WITH BufferUsage AS
    (
        SELECT database_id, COUNT_BIG(*)*8.0/1024 AS CachedMB
        FROM sys.dm_os_buffer_descriptors
        WHERE database_id <> 32767
        GROUP BY database_id
    )
    SELECT ROW_NUMBER() OVER(ORDER BY CachedMB DESC) AS BufferPoolRank,
           DB_NAME(database_id) AS DatabaseName,
           CONVERT(decimal(18,2), CachedMB) AS CachedMB,
           CONVERT(decimal(9,2), CachedMB*100.0/NULLIF(SUM(CachedMB) OVER(),0)) AS BufferPoolPct
    INTO #D17_BufferPool
    FROM BufferUsage;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'17', N'Buffer pool por database', N'#D17_BufferPool', N'BufferPoolRank', N'Distribuição das páginas em cache no momento da coleta.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'17 - Buffer pool', ERROR_MESSAGE());
END CATCH;

INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
VALUES (N'17', N'Uso do log de transações', N'#LogSpace', N'LogSpaceUsedPct DESC', N'DBCC SQLPERF(LOGSPACE).');

IF OBJECT_ID('tempdb..#D17_VLF') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'17', N'VLFs por arquivo de log', N'#D17_VLF', N'VLFCount DESC', N'Referência: acima de 1.000 VLFs merece revisão do crescimento do log.');

/* DETALHAMENTO 18 - WAIT STATS (V9) */
BEGIN TRY
    SELECT TOP (20)
           wait_type,
           waiting_tasks_count,
           CONVERT(decimal(18,2), wait_time_ms/1000.0) AS WaitTimeSec,
           CONVERT(decimal(18,2), signal_wait_time_ms/1000.0) AS SignalWaitSec,
           CONVERT(decimal(18,2), wait_time_ms*1.0/NULLIF(waiting_tasks_count,0)) AS AvgWaitMs,
           CONVERT(decimal(9,2), wait_time_ms*100.0/NULLIF(SUM(wait_time_ms) OVER(),0)) AS PctTotal
    INTO #D18_Waits
    FROM sys.dm_os_wait_stats
    WHERE waiting_tasks_count > 0
      AND wait_type NOT LIKE N'SLEEP[_]%'
      AND wait_type NOT LIKE N'PREEMPTIVE[_]XE[_]%'
      AND wait_type NOT LIKE N'QDS[_]%'
      AND wait_type NOT LIKE N'PARALLEL[_]REDO[_]%'
      AND wait_type NOT IN
      (
        N'BROKER_EVENTHANDLER', N'BROKER_RECEIVE_WAITFOR', N'BROKER_TASK_STOP', N'BROKER_TO_FLUSH',
        N'BROKER_TRANSMITTER', N'CHECKPOINT_QUEUE', N'CHKPT', N'CLR_AUTO_EVENT', N'CLR_MANUAL_EVENT',
        N'CLR_SEMAPHORE', N'CXCONSUMER', N'DBMIRROR_DBM_EVENT', N'DBMIRROR_EVENTS_QUEUE',
        N'DBMIRROR_WORKER_QUEUE', N'DBMIRRORING_CMD', N'DIRTY_PAGE_POLL', N'DISPATCHER_QUEUE_SEMAPHORE',
        N'EXECSYNC', N'FSAGENT', N'FT_IFTS_SCHEDULER_IDLE_WAIT', N'FT_IFTSHC_MUTEX',
        N'HADR_CLUSAPI_CALL', N'HADR_FILESTREAM_IOMGR_IOCOMPLETION', N'HADR_LOGCAPTURE_WAIT',
        N'HADR_NOTIFICATION_DEQUEUE', N'HADR_TIMER_TASK', N'HADR_WORK_QUEUE', N'KSOURCE_WAKEUP',
        N'LAZYWRITER_SLEEP', N'LOGMGR_QUEUE', N'MEMORY_ALLOCATION_EXT', N'ONDEMAND_TASK_QUEUE',
        N'PWAIT_ALL_COMPONENTS_INITIALIZED', N'PWAIT_DIRECTLOGCONSUMER_GETNEXT',
        N'PWAIT_EXTENSIBILITY_CLEANUP_TASK', N'REDO_THREAD_PENDING_WORK', N'REQUEST_FOR_DEADLOCK_SEARCH',
        N'RESOURCE_QUEUE', N'SERVER_IDLE_CHECK', N'SNI_HTTP_ACCEPT', N'SOS_WORK_DISPATCHER',
        N'SP_SERVER_DIAGNOSTICS_SLEEP', N'SQLTRACE_BUFFER_FLUSH', N'SQLTRACE_INCREMENTAL_FLUSH_SLEEP',
        N'SQLTRACE_WAIT_ENTRIES', N'VDI_CLIENT_OTHER', N'WAIT_FOR_RESULTS', N'WAITFOR',
        N'WAITFOR_TASKSHUTDOWN', N'WAIT_XTP_RECOVERY', N'WAIT_XTP_HOST_WAIT', N'WAIT_XTP_OFFLINE_CKPT_NEW_LOG',
        N'WAIT_XTP_CKPT_CLOSE', N'XE_DISPATCHER_JOIN', N'XE_DISPATCHER_WAIT', N'XE_TIMER_EVENT',
        N'XE_LIVE_TARGET_TVF', N'UCS_SESSION_REGISTRATION', N'BROKER_DISPATCHER'
      )
    ORDER BY wait_time_ms DESC;
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'18', N'Top 20 wait stats', N'#D18_Waits', N'WaitTimeSec DESC', N'Esperas acumuladas desde o último restart (ou limpeza), excluindo esperas benignas/ociosas. PctTotal é relativo às esperas filtradas.');
END TRY
BEGIN CATCH
    INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'18 - Wait stats', ERROR_MESSAGE());
END CATCH;


/* ============================================================================
   DETALHAMENTOS 19 a 24 (V10) - registro das seções coletadas no bloco V10
============================================================================ */
IF OBJECT_ID('tempdb..#D19_PoliticaSenha') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'19', N'Política de senha dos logins SQL', N'#D19_PoliticaSenha', N'PolDisabled, PolBlankPassword DESC, PolPasswordEqualsLogin DESC, PolPolicyChecked, PolLogin',
            N'CHECK_POLICY/CHECK_EXPIRATION, última troca de senha, bloqueio e verificação de senha em branco ou igual ao login (PWDCOMPARE).');

IF OBJECT_ID('tempdb..#D19_ServerRoles') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'19', N'Membros de server roles', N'#D19_ServerRoles', N'SrvRoleName, SrvRoleMember',
            N'Todas as server roles (fixas e criadas pelo usuário). securityadmin deve ser tratada como sysadmin.');

IF OBJECT_ID('tempdb..#D19_Permissoes') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'19', N'Permissões explícitas de servidor', N'#D19_Permissoes', N'PermPrincipal, PermName',
            N'sys.server_permissions, omitindo o GRANT CONNECT SQL padrão e as permissões padrão do role public.');

IF OBJECT_ID('tempdb..#D20_Objetos') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'20', N'Objetos por database', N'#D20_Objetos', N'ObjDatabase',
            N'Objetos de usuário (is_ms_shipped = 0): tabelas, views, procedures, functions, triggers, índices, sinônimos e schemas.');

IF @ColetarFragmentacao = 1 AND OBJECT_ID('tempdb..#D20_FragResumo') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'20', N'Resumo de fragmentação por database', N'#D20_FragResumo', N'FragResCriticos DESC, FragResDatabase',
            N'Índices com ' + CONVERT(nvarchar(20), @FragMinPaginas) + N' páginas ou mais, modo LIMITED. Críticos >= 70%, moderados 30-70%.');

IF @ColetarFragmentacao = 1 AND OBJECT_ID('tempdb..#D20_Fragmentacao') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'20', N'Índices fragmentados (>= 30%)', N'#D20_Fragmentacao', N'FragPct DESC, FragPages DESC',
            N'Até 500 índices, do mais fragmentado para o menos. FragIndexId 0 = heap (ALTER TABLE ... REBUILD).');

IF @ColetarFragmentacao = 0
    INSERT INTO #Secoes (Grupo, Titulo, Nota)
    VALUES (N'20', N'Índices fragmentados', N'Coleta de fragmentação desligada nesta execução (@ColetarFragmentacao = 0).');

IF OBJECT_ID('tempdb..#D20_MissingIndexes') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'20', N'Missing indexes (top 25)', N'#D20_MissingIndexes', N'MIImprovement DESC',
            N'Sugestões do otimizador desde o último restart. Não consideram índices existentes nem o custo de escrita: consolidar antes de criar.',
            N'MIEquality,MIInequality,MIIncluded,MICreateStatement');

IF OBJECT_ID('tempdb..#D21_TopDuracao') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'21', N'Top 20 queries por duração média', N'#D21_TopDuracao', N'TQAvgDurationMs DESC',
            N'sys.dm_exec_query_stats: somente planos em cache, com ao menos 2 execuções. Tempos em milissegundos.', N'TQSqlText');

IF OBJECT_ID('tempdb..#D21_TopCPU') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'21', N'Top 20 queries por CPU total', N'#D21_TopCPU', N'TCTotalCpuMs DESC',
            N'Consumo acumulado de CPU desde a compilação do plano. Tempos em milissegundos.', N'TCSqlText');

IF OBJECT_ID('tempdb..#D22_Sessoes') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'22', N'Sessões de usuário no momento da coleta', N'#D22_Sessoes', N'SessBlockedBy DESC, SessOpenTran DESC, SessCpuMs DESC',
            N'Fotografia do momento da execução do assessment (exclui a própria sessão). SessBlockedBy <> 0 = sessão bloqueada.', N'SessSqlText');

IF OBJECT_ID('tempdb..#D23_BackupHist') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'23', N'Histórico de backups Full/Diferencial (7 dias)', N'#D23_BackupHist', N'BkStart DESC',
            N'msdb desta instância. BkDeviceType 2 = disco, 7 = dispositivo virtual (ferramenta externa), 9 = URL. BkCompressionRatio ~1 = sem compressão.', N'BkDevice');

IF OBJECT_ID('tempdb..#D23_BackupLog') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'23', N'Backups de log - resumo diário (7 dias)', N'#D23_BackupLog', N'LogBkDatabase, LogBkDay DESC',
            N'Quantidade, volume e janela dos backups de log por database e dia.');

IF OBJECT_ID('tempdb..#D24_JobAgenda') IS NOT NULL
    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'24', N'Agendamentos dos jobs', N'#D24_JobAgenda', N'SchJobName, SchScheduleName',
            N'Frequência, dias, horário e próxima execução de cada agendamento (exceto assinaturas do Reporting Services).');

/* ============================================================================
   DETALHAMENTO 25 (V11) - LOG SHIPPING (exibido na guia Alta Disponibilidade)
============================================================================ */
IF @K_LSPrim + @K_LSSec + @K_LSMon = 0
    INSERT INTO #Secoes (Grupo, Titulo, Nota)
    VALUES (N'25', N'Log shipping', N'Log shipping não configurado nesta instância (nenhum registro em msdb.dbo.log_shipping_*).');
ELSE
BEGIN
    IF EXISTS (SELECT 1 FROM #D25_LSPrimario)
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'25', N'Log shipping - databases primárias', N'#D25_LSPrimario', N'LSPDatabase',
                N'Configuração do envio de logs: pasta local e compartilhamento de backup, retenção (min), compressão, servidor monitor, secundários, último backup de log e minutos decorridos x limite de alerta.');

    IF EXISTS (SELECT 1 FROM #D25_LSSecundario)
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'25', N'Log shipping - databases secundárias', N'#D25_LSSecundario', N'LSSDatabase',
                N'Origem (primário), pastas de cópia, modo de restore (NORECOVERY ou STANDBY), atraso configurado, última cópia, último restore, latência e limite de alerta. Limite efetivo = LSSLimiteMin + LSSAtrasoRestoreMin.');

    IF EXISTS (SELECT 1 FROM #D25_LSMonitor)
        INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
        VALUES (N'25', N'Log shipping - databases monitoradas (servidor monitor)', N'#D25_LSMonitor', N'LSMStatus, LSMServidor, LSMDatabase',
                N'Databases de outras instâncias que usam esta instância como servidor monitor. Tempos calculados em UTC.');

    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao)
    VALUES (N'25', N'Log shipping - jobs', N'#D25_LSJobs', N'LSJFuncao, LSJJob',
            N'Jobs LSBackup (primário), LSCopy e LSRestore (secundário) e LSAlert (monitor): habilitado, última execução, resultado e próxima execução.');

    INSERT INTO #Secoes (Grupo, Titulo, Tabela, OrdenarPor, Descricao, ColunasLargas)
    VALUES (N'25', N'Log shipping - erros recentes (7 dias)', N'#D25_LSErros', N'LSEData DESC',
            N'Últimas 50 mensagens de msdb.dbo.log_shipping_monitor_error_detail.', N'LSEMensagem');
END;


END; /* IF @Modo IN (1,3) */

/* ============================================================================
   EXIBIÇÃO DOS DETALHAMENTOS (SSMS) + MONTAGEM DO RELATÓRIO HTML
============================================================================ */
DECLARE @Html nvarchar(max) = N'', @Css nvarchar(max) = N'', @ImgFooter nvarchar(max) = N'';
DECLARE @Det nvarchar(max) = N'', @Chips nvarchar(max) = N'', @Tmp nvarchar(max), @Tmp2 nvarchar(max);
DECLARE @sOrdem int, @sGrupo nvarchar(10), @sTitulo nvarchar(200), @sTabela sysname, @sOrd nvarchar(400),
        @sNota nvarchar(2000), @sDesc nvarchar(1000), @sLargas nvarchar(400);
DECLARE @rHdr nvarchar(max), @rApply nvarchar(max), @rCells nvarchar(max), @rKV nvarchar(max),
        @rSql nvarchar(max), @rRows nvarchar(max), @rTotal int, @rNCols int, @rVertical bit;
DECLARE @LstOk nvarchar(1000)   = N'N''OK'',N''HEALTHY'',N''SYNCHRONIZED'',N''ONLINE'',N''CONNECTED'',N''STARTED'',N''Running'',N''Sucesso''';
DECLARE @LstWarn nvarchar(1000) = N'N''ATENÇÃO'',N''PARTIALLY_HEALTHY'',N''RESTORING'',N''RECOVERING'',N''Stopped'',N''Paused'',N''Cancelado'',N''Nova tentativa''';
DECLARE @DetSec nvarchar(max);
CREATE TABLE #V11_HtmlDet (Ordem int NOT NULL, Grupo nvarchar(10) NOT NULL, Titulo nvarchar(200) NOT NULL,
                           Guia varchar(20) NULL, Html nvarchar(max) NULL);
DECLARE @LstCrit nvarchar(1000) = N'N''CRÍTICO'',N''NOT_HEALTHY'',N''SUSPECT'',N''EMERGENCY'',N''OFFLINE'',N''RECOVERY_PENDING'',N''DISCONNECTED'',N''NOT SYNCHRONIZING'',N''NOT_SYNCHRONIZING'',N''Falha''';

DECLARE curSecoes CURSOR LOCAL FAST_FORWARD FOR
    SELECT Ordem, Grupo, Titulo, Tabela, ISNULL(OrdenarPor, N'1'), Nota, Descricao, ISNULL(ColunasLargas, N'')
    FROM #Secoes
    ORDER BY Ordem;

OPEN curSecoes;
FETCH NEXT FROM curSecoes INTO @sOrdem, @sGrupo, @sTitulo, @sTabela, @sOrd, @sNota, @sDesc, @sLargas;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF @sTabela IS NULL
    BEGIN
        /* seção apenas informativa */
        IF @SomenteHTML = 0
            SELECT @sNota AS Information;

        IF @GerarHTML = 1
        BEGIN
            SET @Chips += N'<a href="#s' + CONVERT(nvarchar(10), @sOrdem) + N'"><b>' + @sGrupo + N'</b>' + REPLACE(REPLACE(REPLACE(@sTitulo, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</a>';
            SET @DetSec = N'<details id="s' + CONVERT(nvarchar(10), @sOrdem) + N'" open="open"><summary><span class="g">' + @sGrupo + N'</span>'
                      + REPLACE(REPLACE(REPLACE(@sTitulo, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'<span class="c">informativo</span></summary><div class="note">'
                      + REPLACE(REPLACE(REPLACE(ISNULL(@sNota, N'-'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</div></details>';
            SET @Det += @DetSec;
            INSERT INTO #V11_HtmlDet (Ordem, Grupo, Titulo, Html) VALUES (@sOrdem, @sGrupo, @sTitulo, @DetSec);
        END;
    END
    ELSE
    BEGIN
        /* resultset no SSMS (mesma ordem da coleta) */
        IF @SomenteHTML = 0
        BEGIN
            SET @rSql = N'SELECT * FROM ' + @sTabela + N' ORDER BY ' + @sOrd + N';';
            BEGIN TRY
                EXEC sys.sp_executesql @rSql;
            END TRY
            BEGIN CATCH
                INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (@sGrupo + N' - ' + @sTitulo + N' (exibição)', ERROR_MESSAGE());
            END CATCH;
        END;

        /* tabela HTML */
        IF @GerarHTML = 1
        BEGIN
            BEGIN TRY
                SELECT @rHdr = NULL, @rApply = NULL, @rCells = NULL, @rKV = NULL, @rRows = NULL, @rTotal = 0, @rNCols = 0;

                SELECT
                    @rNCols = COUNT(*),
                    @rHdr   = STRING_AGG(CAST(N'<th>' + REPLACE(REPLACE(REPLACE(c.name, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</th>' AS nvarchar(max)), N'')
                              WITHIN GROUP (ORDER BY c.column_id),
                    @rApply = STRING_AGG(CAST(N'v' + CONVERT(nvarchar(10), c.column_id)
                              + N' = REPLACE(REPLACE(REPLACE(ISNULL(' + ve.expr
                              + N', N''''), N''&'', N''&amp;''), N''<'', N''&lt;''), N''>'', N''&gt;'')' AS nvarchar(max)), N', ')
                              WITHIN GROUP (ORDER BY c.column_id),
                    @rCells = STRING_AGG(CAST(ce.expr AS nvarchar(max)), N' + ')
                              WITHIN GROUP (ORDER BY c.column_id),
                    @rKV    = STRING_AGG(CAST(N'(' + CONVERT(nvarchar(10), c.column_id) + N', N'''
                              + REPLACE(REPLACE(REPLACE(REPLACE(c.name, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N'''', N'''''') + N''', x.v' + CONVERT(nvarchar(10), c.column_id) + N')' AS nvarchar(max)), N', ')
                              WITHIN GROUP (ORDER BY c.column_id)
                FROM tempdb.sys.columns AS c
                CROSS APPLY (SELECT N't.' + QUOTENAME(c.name) AS col,
                                    CONVERT(nvarchar(10), c.column_id) AS id,
                                    CASE WHEN c.system_type_id IN (48,52,56,127,59,60,62,106,108,122) THEN 1 ELSE 0 END AS isnum,
                                    CASE WHEN (N',' + @sLargas + N',') LIKE (N'%,' + c.name + N',%') THEN N' class="w"' ELSE N'' END AS wcls) AS q
                CROSS APPLY (SELECT CASE
                                WHEN c.system_type_id = 40 THEN N'CONVERT(nvarchar(10), ' + q.col + N', 103)'
                                WHEN c.system_type_id IN (42,43,58,61) THEN N'CONVERT(nvarchar(10), ' + q.col + N', 103) + N'' '' + CONVERT(nvarchar(8), ' + q.col + N', 108)'
                                WHEN c.system_type_id = 104 THEN N'CASE ' + q.col + N' WHEN 1 THEN N''Sim'' WHEN 0 THEN N''Não'' END'
                                WHEN c.system_type_id IN (59,60,62,106,108,122) THEN N'FORMAT(' + q.col + N', N''N2'', N''pt-BR'')'
                                WHEN c.system_type_id IN (165,173) THEN N'CONVERT(nvarchar(max), ' + q.col + N', 1)'
                                WHEN c.system_type_id = 98 THEN N'CONVERT(nvarchar(4000), ' + q.col + N')'
                                ELSE N'CONVERT(nvarchar(max), ' + q.col + N')'
                             END AS expr) AS ve
                CROSS APPLY (SELECT CASE WHEN q.isnum = 1
                                THEN N'N''<td class="n">'' + ISNULL(NULLIF(x.v' + q.id + N', N''''), N''—'') + N''</td>'''
                                ELSE N'CASE WHEN x.v' + q.id + N' IN (' + @LstOk + N') THEN N''<td><span class="b ok">'' + x.v' + q.id + N' + N''</span></td>'''
                                   + N' WHEN x.v' + q.id + N' IN (' + @LstWarn + N') THEN N''<td><span class="b warn">'' + x.v' + q.id + N' + N''</span></td>'''
                                   + N' WHEN x.v' + q.id + N' IN (' + @LstCrit + N') THEN N''<td><span class="b crit">'' + x.v' + q.id + N' + N''</span></td>'''
                                   + N' ELSE N''<td' + q.wcls + N'>'' + ISNULL(NULLIF(x.v' + q.id + N', N''''), N''—'') + N''</td>'' END'
                             END AS expr) AS ce
                WHERE c.object_id = OBJECT_ID(N'tempdb..' + @sTabela);

                SET @rSql = N'SELECT @Total = COUNT(*) FROM ' + @sTabela + N';';
                EXEC sys.sp_executesql @rSql, N'@Total int OUTPUT', @Total = @rTotal OUTPUT;

                /* tabela com 1 linha e muitas colunas: exibe na vertical (propriedade | valor) */
                SET @rVertical = CASE WHEN @rTotal = 1 AND @rNCols > 6 THEN 1 ELSE 0 END;

                IF @rVertical = 1
                    SET @rSql = N'SELECT @Rows = STRING_AGG(CAST(N''<tr><th>'' + kv.k + N''</th><td>'' + ISNULL(NULLIF(kv.v, N''''), N''—'') + N''</td></tr>'' AS nvarchar(max)), N'''') WITHIN GROUP (ORDER BY kv.o)'
                              + NCHAR(13) + NCHAR(10) + N'FROM (SELECT TOP (1) * FROM ' + @sTabela + N') AS t'
                              + NCHAR(13) + NCHAR(10) + N'CROSS APPLY (SELECT ' + @rApply + N') AS x'
                              + NCHAR(13) + NCHAR(10) + N'CROSS APPLY (VALUES ' + @rKV + N') AS kv(o, k, v);';
                ELSE
                    SET @rSql = N'SELECT @Rows = STRING_AGG(CAST(N''<tr>'' + ' + @rCells + N' + N''</tr>'' AS nvarchar(max)), N'''') WITHIN GROUP (ORDER BY ' + @sOrd + N')'
                              + NCHAR(13) + NCHAR(10) + N'FROM (SELECT TOP (@Max) * FROM ' + @sTabela + N' ORDER BY ' + @sOrd + N') AS t'
                              + NCHAR(13) + NCHAR(10) + N'CROSS APPLY (SELECT ' + @rApply + N') AS x;';

                IF @rTotal > 0
                    EXEC sys.sp_executesql @rSql, N'@Max int, @Rows nvarchar(max) OUTPUT', @Max = @HtmlMaxLinhas, @Rows = @rRows OUTPUT;

                SET @Chips += N'<a href="#s' + CONVERT(nvarchar(10), @sOrdem) + N'"><b>' + @sGrupo + N'</b>' + REPLACE(REPLACE(REPLACE(@sTitulo, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</a>';

                SET @Tmp = N'<details id="s' + CONVERT(nvarchar(10), @sOrdem) + N'" open="open"><summary><span class="g">' + @sGrupo + N'</span>'
                         + REPLACE(REPLACE(REPLACE(@sTitulo, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'<span class="c">'
                         + CASE @rTotal WHEN 0 THEN N'nenhum registro' WHEN 1 THEN N'1 registro' ELSE CONVERT(nvarchar(20), @rTotal) + N' registros' END
                         + N'</span></summary>'
                         + CASE WHEN @sDesc IS NULL THEN N'' ELSE N'<p class="desc">' + REPLACE(REPLACE(REPLACE(@sDesc, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</p>' END;

                IF @rVertical = 1
                    SET @Tmp += N'<div class="tw"><table class="data kv"><thead><tr><th>Propriedade</th><th>Valor</th></tr></thead><tbody>'
                              + ISNULL(@rRows, N'') + N'</tbody></table></div>';
                ELSE
                    SET @Tmp += N'<div class="tw"><table class="data"><thead><tr>' + ISNULL(@rHdr, N'<th>-</th>') + N'</tr></thead><tbody>'
                              + ISNULL(@rRows, N'<tr><td class="empty" colspan="' + CONVERT(nvarchar(10), @rNCols) + N'">Nenhum registro encontrado.</td></tr>')
                              + N'</tbody></table></div>';

                IF @rTotal > @HtmlMaxLinhas
                    SET @Tmp += N'<div class="note">Exibindo ' + CONVERT(nvarchar(20), @HtmlMaxLinhas) + N' de ' + CONVERT(nvarchar(20), @rTotal)
                              + N' registros. A lista completa está no resultset correspondente (ou aumente @HtmlMaxLinhas).</div>';

                SET @DetSec = @Tmp + N'</details>';
                SET @Det += @DetSec;
                INSERT INTO #V11_HtmlDet (Ordem, Grupo, Titulo, Html) VALUES (@sOrdem, @sGrupo, @sTitulo, @DetSec);
            END TRY
            BEGIN CATCH
                INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (@sGrupo + N' - ' + @sTitulo + N' (HTML)', ERROR_MESSAGE());
            END CATCH;
        END;
    END;

    FETCH NEXT FROM curSecoes INTO @sOrdem, @sGrupo, @sTitulo, @sTabela, @sOrd, @sNota, @sDesc, @sLargas;
END;

CLOSE curSecoes;
DEALLOCATE curSecoes;


IF @GerarHTML = 1
BEGIN
    /* ------------------------------------------------------------------------
       Estilo e imagens do papel timbrado (embutidos, o HTML é autocontido)
    ------------------------------------------------------------------------ */
    SET @Css += N':root{--red:#D32225;--red-dk:#A72A2B;--ink:#2F2F2E;--ink2:#3C3C3B;--muted:#76766F;--line:#E4E4E1;--soft:#F6F6F4;--bg:#EDEDEA;--ok:#2E7D4F;--ok-bg:#E6F2EA;--warn:#A86400;--warn-bg:#FFF2DB;--crit:#B3121B;--crit-bg:#FBE3E4;--info:#4F6272;--info-bg:#E9EEF2}';
    SET @Css += N'*{box-sizing:border-box}';
    SET @Css += N'html{background:var(--bg)}';
    SET @Css += N'body{margin:0;padding:32px 16px;background:var(--bg);color:var(--ink);font:14px/1.5 "Lato","Segoe UI",Roboto,Helvetica,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}';
    SET @Css += N'.page{position:relative;max-width:1180px;margin:0 auto;background:#fff url(data:image/png;base64,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';
    SET @Css += N'ngpmkrdtmaUIHEEnBhFcJ5nAlmDPBPJBICiYIpkgKJgimSL5ZME/iSWsQTJEUTBBMkRRMEEyRFEwQTJEUTBBMkdzLDS6YIJgiKZggmCIpmCCYIimYIJgiKZggmCIpmCCYxxLMo4ykYIJgiqRggmCKpGCCYIqkYIJgiqRggmCKpGACexZMkRRMEEyRFEwQTJEUTBBMkRRMEMy3DKZICibsazCrsv+MXzOYIimYsO/7z8lKLHceTJEUTBBMkRRMEMyXBVMkBRMEUyQFE3hZMEVSMEEwe4IpkoIJ9ARTJAUT6AmmSAom0EMkjy+Yg3z/aNdAMEEkEUwQSQQTRBLBBJFEMEEkEUwQSQQTDiqS02p0muQsyXJl0S+MCcHk6CNZ3mH+d5JJkjbfP5ozL98TTASTo45kleSibpv/Vhb86mcZqyQLwUQwOeZIntdtc9Oz6Fc//C2YCCYiKZgIJiL5gkgKJoKJSL5s0QsmgolICiaCiUjudtELJoKJSAomgolICiaCiUj+kUUvmAgmIimYCCYiKZgIJiIpmAgmIimYCCYiuY+RFEwEE5EUTAQTkRRMBBORFEwEE5EUTAQTkRRMEEyRPNpFL5gIJiIpmPyJYCZp67ZZmpBICiZ0B3O4EsyZYIqkYIJgiqRgCiaCKZIIJoIpkiL5ysFs67aZmw6CKZJsD+ayLPiJYCKYIolgIpgiiWAimCKJYCKYIimSgolgIpKCiWAKpkgKJgimSAqmYCKYIolgIpgiiWAimCKJYCKY+xvJYZKrum2+WCKCCYK5FskyjH/qtvnHshBMEEyRRDARTJFEMBHMXUbyc5Lbum1am10wBRPB3IzkdZI7kUQwEczuSF4kWdZtc2fzIpgI5mYkz5JUddt8tUkRTARzM5LeK4lgIpjbIln+wv8k+eIOfwgmgtkdyavyF5rYZAgmxx7Mrkh6XhLBRDB7IjlI8rlum39tGgSTYw/mRiTLX+Jzknun3Agmxx7MbZF0yo1gIpg9kTxJ8ne8ys2fD2ZVFn4Ek19YOycr/9iOfyeYnZEsf8hlktRt883IeQeLfliOEASTNw1mXyQHST6Xo0n3ukAwOcpgbo3kytHksm6bW+NFMDnGYP4skoMkf8VzkwgmRxrM3kiWX3yRZOCVbgSTYwzmcyJ5Uo4mv9VtMzNCBJNjCuZPI1l+4jjJZbyIg2ByZJ4VybK4rpLEaTeCiUhuP+1++rjivdEhmIhk90L6nOQ/iwfBRCS7F9BpkoskNxYNgolIdi+ci7Jo/vNCDoIpmCLZvWAu83iljRuhRDAFUyT7Q/nVJ3IQTMEUye2hHMdzlCCYIrl1YZzm8cWcW1czB8EUye4FUSW5SjJJcud5ShBMkdxcCIM8fnzxJI+f9bYAQDBFsmMRnCU5d1QJgimSPz+qHJZQ+igjCKZIdmz4qhxVDpLc5Q1uJg6Cyd5EsiOWw3Iafu+9lSCYItm9wc/KBl+UYM4EEwRTJDc3+Oo9JQQTBFMkBRMEUyQFE95bMGd127SmcyCRFEx4lWCerATTTf0OJZKCCYIpkoIJgimSggmCKZKCCYIpkoIJgimSggkI5vFEUjBBMEVSMEEwRVIwQTBFUjBBMEVSMEEwRVIwgT0OpkgKJgimSAomCKZICiYIpkgKJgjmWwZTJAUTBFMk9yKYbdyTBN5dMEXy/QVzGTdxgncTTJEUTBBMkRRMEMyXBVMkBRMEUyQFE3hZMEVSMEEwe4IpkoIJ9ARTJAUT6A7mOMlYJAUT6CGSgimYIJIIJogkggkiiWCCSCKYIJIIJogkggkiiWAKJiIJgolIgmCCSCKYIJIIJogkggkiiWCaDiIJgolIgmAikvAWwazKtwQTkYQtwVy9J4lgIpIgmIgkCCYiCYKJSIJgIpIgmIgkCCYiCQgmIgmCiUiCYCKSIJiIJAgmIgmCiUiCYAqmSAKCKZKAYIokIJiIJAgmIgmCiUiCYCKSIJiCKZKAYIokIJgiCQimSAKCKZLAwQVzVrdNazoiCWwP5slKMGemI5KAYIokIJgiCQimSAKCKZKAYIokIJgiCQjmgQVTJAHBFElAMEUSEEyRBARTJAHBfCfBFElAMEUSEEyRBARTJAHBfMtgiiQgmCIJCKZIAuw8mCIJCKZIArwsmCIJCGZPMEUSEMyeYIokQE8wRRKgJ5giCdDjoxEAiCSASAKIJIBIAogkgEgCiCSASAKIJAAiCSCSACIJ8Jo+GQHsj2k1uk5SJZknmeTxwrALk3k9LpUG+xPI0yQXSW6SDPJ4UdhxkoVgiiSIZDU6T5K6be7Wvj8WTKfbwBblxlWztWCeT6uRYIokIJgiCQimSAKCKZLAew/m0/2m5yb1yKvbsCe2vbq9w99/9VXyZQnm5NiDKZIgkoIpkiCSgimSIJKCKZIgkn8ukscYTJEEkRRMkQSRFEyRBJEUTJEEkdyfSB5CMEUSRFIwRRJEUjBFEkRSMEUSRPJwI/kegymSIJKCKZIgkoIpkiCSgrnzYIokiKRgiiSIpGC+LJgiCSIpmCIJIsnLgimSIJKC2RNMkQSRNNueYIokiCQ9wXTfbYAV6/cl/2gkANuDKZIAPUQSQCQBRBJAJAFEEkAkAUQSQCQBRBJAJAEQSQCRBBBJAJEEEEkAkQQQSQCRBBBJAJEEQCQBRBJAJAFEEkAkAUQSQCQBRBJAJAEQSQCRBBBJAJEEEEmg1zLJwBhEEujWJhlPq5FQvqEPDw8PpgB7YlqNLpKcJrlLMqvbZmEqIgn8GMqqhHKcZJFkIpgiCXQHc1xiKZgiCQimSAKCKZKAYIokIJgiCQimSAKCKZIAgimSgGCKJCCYIgkIpkgCexjMWQnmXCQBtgdzWYI52edgiiQgmCIJCKZIAoIpkoBgiiQgmO8kmCIJCKZIAoIpkoBgiiTAWwZTJAHBFElAMF8WTJEEBFMkAcF8WTBFEhDMnmCKJCCYPcEUSYCeYIokQE8wRRKgx0cjABBJAJEEEEkAkQQQSQCRBBBJAJE';
    SET @Css += N'EEEkARBJAJAFEEkAkAf6QT0YAsGlajU6SjEUSYC2MK4+5i+4CwrgWxiSTJLO6bRYiCQjjWhhXf65IAsK4FkaRBISxJ4wiCQjjM4kkIIwiCQijSALCuHMiCQijSALCKJKAMIokIIwiCQjjOyGSgDCKJCCMIgkIo0gCwiiSgDCKJCCMIgkcVhhnSdpDD6NIAr8axlkJ4/LY5iGSgDCKJCCMIgkIo0gCwiiSgDCKJCCMIgkIo0gCwiiSwHsI5GmSi5RPvQjj6/tkBLA3gRyUQH6r22ZmIm/joxHA3qiStAIpkkC3QR6vvINIAogkgEgCiCSASAKIJAAiCSCSACIJIJIAIgkgkgAiCSCSACIJIJIAiCSASAKIJIBIAogkgEgCiCSASAKIJAAiCSCSACIJIJIAIgkgkgAiCSCSACIJIJIArJtWo5NPxgDwYxiTjJ8eIgkI41oYk8yTTJLciiQgjGthrNtm8fTzRBIQxrUwrhJJ4NjC2CaZ9YVRJIFjC+OsPL7VbbP8ld9LJAFhFElAGEUSEMblrv88kQSEUSQBYRRJQBhFEhBGkQSE8Z0QSUAYRRIQRpEEhFEkAWEUSUAYRRIQRpEEhFEkAWEUSUAYj5BIgjAKo0gCwiiSgDCKJCCMIgkIo0gCwiiSwO4sk5xPq9F9+VoY38CHh4cHU4D9OWL8XL4crIRxJowiCXyPZZVkLowiCfDHfTQCgO3+B+8VT1UP2uBdAAAAAElFTkSuQmCC) no-repeat right 0 top 28px;background-size:190px auto;box-shadow:0 2px 24px rgba(0,0,0,.08);padding:56px 64px 28px 76px;overflow:hidden}';
    SET @Css += N'.page::before{content:"";position:absolute;left:0;top:92px;width:30px;height:168px;background:url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADoAAAE+CAYAAADcXQOKAAADLUlEQVR42u3awW2jUBDG8c+R70kHDpflwCEuwR04JTgVrDvYlEAHSzogFYR0gKXlwAnTAa6APSxICDmJLewsvPmP5IMPHH4eZt578zyr61oW4kZG4kaSMs9fZp5/7zJ09uf+Ryxp3XzfSUraT1DklUvQz4r0vYNOXIZ249DLduoqtB9lD753FdqPUdf3JaGjru9rQkdV398FPQaPv7O+/xf0WGOLJUXXyvZYoH102KArl6Hd1zuUFF4CPGboRcFTOL3cSvolKc08f2XhmLaQ9JZ5fph5/p2F8+hPSUnm+UsLB++HBrtyHdrW7lvm+RvXoW38PgXryszoS6xLw7HwswblEvRWUvzR0uPauHMhKbIAlaT1sWXH1QF2ZAW66Hdhl68ktlagD93lxvVLpo0V6KMV6KK9JbRwP7qyAjWT0aUV6J0VqIACBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoU6PCorEBTK9C9FWhiAVoGRW4io4mVrhtagJZBkacWoM8WNgxlUOSRBejWwhbwNSjy2HXoQdLGwqZ+FRR55Tr0qbucuAp96ndZF6FfIiVpPvHG8xgUeeLyhGHXNJ7k1AemCH1pkOk5D03p1S0lbY9tBlyBHprjVvjRGukC9EXSczslGBLzkb6i0dAMjhVaSoolRec2mbFDD/o3z4klJZd4NccEfe/A0u/+Za8J3bVZO2dhnwK0rbOkyVqlEcV8YJ11YXuNOM6FvnZgqSYU8xPqrG0giSYc8yZL6ynU2ZCY1XWtzPOXkqqx19lgqIUw8x+GvxtS+i0JqTS8AAAAAElFTkSuQmCC) no-repeat left top;background-size:100% 100%}';
    SET @Css += N'a{color:inherit}';
    SET @Css += N'h1,h2,h3{font-weight:700;color:var(--ink2);margin:0}';
    SET @Css += N'.cover{position:relative;padding:8px 210px 30px 0;border-bottom:1px solid var(--line)}';
    SET @Css += N'.kicker{font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:var(--red);font-weight:700}';
    SET @Css += N'.cover h1{font-size:30px;line-height:1.2;margin:10px 0 4px;letter-spacing:-.01em}';
    SET @Css += N'.client{font-size:16px;color:var(--muted)}';
    SET @Css += N'.client b{color:var(--ink2)}';
    SET @Css += N'.server{margin:22px 0 0;display:flex;flex-direction:column}';
    SET @Css += N'.server .lbl{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted)}';
    SET @Css += N'.server .val{font-size:26px;font-weight:700;color:var(--ink2);font-variant-numeric:tabular-nums;word-break:break-all}';
    SET @Css += N'.meta{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px 24px;margin-top:22px}';
    SET @Css += N'.meta div{border-left:3px solid var(--red);padding:2px 0 2px 10px}';
    SET @Css += N'.meta span{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}';
    SET @Css += N'.meta b{display:block;font-size:14px;color:var(--ink2);word-break:break-word}';
    SET @Css += N'.toc{display:flex;flex-wrap:wrap;gap:6px 18px;padding:14px 0;border-bottom:1px solid var(--line);font-size:13px}';
    SET @Css += N'.toc a{text-decoration:none;color:var(--muted)}';
    SET @Css += N'.toc a:hover{color:var(--red)}';
    SET @Css += N'.toc a b{color:var(--red);margin-right:6px}';
    SET @Css += N'section{margin-top:40px}';
    SET @Css += N'h2{display:flex;align-items:baseline;gap:12px;font-size:21px;margin-bottom:16px;padding-bottom:8px;border-bottom:2px solid var(--ink2)}';
    SET @Css += N'h2 .num{color:var(--red);font-size:15px;letter-spacing:.06em}';
    SET @Css += N'h3{font-size:15px;margin-bottom:12px}';
    SET @Css += N'.lead{color:var(--muted);margin:-6px 0 16px;max-width:820px}';
    SET @Css += N'.exec{display:grid;grid-template-columns:230px 1fr;gap:24px;align-items:stretch}';
    SET @Css += N'.gauge{display:flex;flex-direction:column;align-items:center;justify-content:center;background:var(--soft);border-radius:10px;padding:18px}';
    SET @Css += N'.gauge svg{width:150px;height:150px}';
    SET @Css += N'.gauge .pct{font-size:30px;font-weight:700;fill:var(--ink2)}';
    SET @Css += N'.gauge .cap{font-size:12px;color:var(--muted);text-align:center;margin-top:6px;max-width:170px}';
    SET @Css += N'.counts{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}';
    SET @Css += N'.cnt{border-radius:10px;padding:16px 18px;background:var(--soft);display:flex;flex-direction:column;justify-content:space-between;border-top:4px solid var(--info)}';
    SET @Css += N'.cnt span{font-size:12px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);font-weight:700}';
    SET @Css += N'.cnt b{font-size:40px;line-height:1.1;color:var(--ink2);font-variant-numeric:tabular-nums}';
    SET @Css += N'.cnt em{font-style:normal;font-size:12px;color:var(--muted)}';
    SET @Css += N'.cnt.crit{border-top-color:var(--crit)}.cnt.crit b{color:var(--crit)}';
    SET @Css += N'.cnt.warn{border-top-color:var(--warn)}.cnt.warn b{color:var(--warn)}';
    SET @Css += N'.cnt.ok{border-top-color:var(--ok)}.cnt.ok b{color:var(--ok)}';
    SET @Css += N'.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:16px}';
    SET @Css += N'.kpi{border:1px solid var(--line);border-radius:10px;padding:14px 16px;background:#fff}';
    SET @Css += N'.kpi .k{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}';
    SET @Css += N'.kpi .v{display:block;font-size:24px;font-weight:700;color:var(--ink2);line-height:1.3;font-variant-numeric:tabular-nums;word-break:break-word}';
    SET @Css += N'.kpi .s{display:block;font-size:12px;color:var(--muted)}';
    SET @Css += N'.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}';
    SET @Css += N'.card{border:1px solid var(--line);border-radius:10px;padding:18px 20px;background:#fff}';
    SET @Css += N'.card+.card{margin-top:0}';
    SET @Css += N'.grid2+.card{margin-top:16px}';
    SET @Css += N'.hb{display:grid;grid-template-columns:minmax(0,34%) 1fr auto;gap:4px 12px;align-items:center;font-size:13px}';
    SET @Css += N'.hb .n{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--ink2)}';
    SET @Css += N'.hb .t{color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap;text-align:right}';
    SET @Css += N'.bar{height:10px;border-radius:5px;background:var(--line);overflow:hidden}';
    SET @Css += N'.bar i{display:block;height:100%;border-radius:5px;background:var(--ink2);font-size:0;line-height:0;color:transparent}';
    SET @Css += N'.bar i.ok{background:var(--ok)}.bar i.warn{background:var(--warn)}.bar i.crit{background:var(--crit)}.bar i.db{background:var(--red)}';
    SET @Css += N'.empty{color:var(--muted);font-style:italic;font-size:13px}';
    SET @Css += N'.cpu svg{width:100%;height:170px;display:block}';
    SET @Css += N'.cpu .grid{stroke:var(--line);stroke-width:1}';
    SET @Css += N'.cpu .ax{font-size:10px;fill:var(--muted)}';
    SET @Css += N'.cpu .l-sql{fill:none;stroke:var(--red);stroke-width:2}';
    SET @Css += N'.cpu .l-oth{fill:none;stroke:#9A9A94;stroke-width:1.5;stroke-dasharray:4 3}';
    SET @Css += N'.legend{display:flex;gap:18px;font-size:12px;color:var(--muted);margin-top:6px}';
    SET @Css += N'.legend span::before{content:"";display:inline-block;width:14px;height:3px;margin-right:6px;vertical-align:middle;background:var(--red)}';
    SET @Css += N'.legend span.o::before{background:#9A9A94}';
    SET @Css += N'.tw{overflow-x:auto;border:1px solid var(--line);border-radius:8px}';
    SET @Css += N'table{width:100%;border-collapse:collapse;font-size:12.5px}';
    SET @Css += N'th{background:var(--ink2);color:#fff;text-align:left;font-weight:700;padding:8px 10px;white-space:nowrap;font-size:11.5px;letter-spacing:.02em}';
    SET @Css += N'td{padding:7px 10px;border-top:1px solid var(--line);vertical-align:top}';
    SET @Css += N'tbody tr:nth-child(even) td{background:#FAFAF8}';
    SET @Css += N'td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}';
    SET @Css += N'table.data td{white-space:nowrap;max-width:520px;overflow:hidden;text-overflow:ellipsis}';
    SET @Css += N'table.data td.w{white-space:normal;min-width:260px}';
    SET @Css += N'tr.grp td{background:var(--soft)!important;font-weight:700;color:var(--ink2);font-size:12px;letter-spacing:.06em;text-transform:uppercase;border-top:1px solid #d6d6d2}';
    SET @Css += N'.b{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:700;letter-spacing:.03em;white-space:nowrap;background:var(--info-bg);color:var(--info)}';
    SET @Css += N'.b.ok{background:var(--ok-bg);color:var(--ok)}';
    SET @Css += N'.b.warn{background:var(--warn-bg);color:var(--warn)}';
    SET @Css += N'.b.crit{background:var(--crit-bg);color:var(--crit)}';
    SET @Css += N'td.obs{color:var(--muted)}';
    SET @Css += N'.chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px}';
    SET @Css += N'.chips a{text-decoration:none;font-size:12px;border:1px solid var(--line);border-radius:14px;padding:3px 10px;color:var(--ink2);background:#fff}';
    SET @Css += N'.chips a:hover{border-color:var(--red);color:var(--red)}';
    SET @Css += N'.chips a b{color:var(--red);margin-right:4px}';
    SET @Css += N'details{border:1px solid var(--line);border-radius:10px;margin-bottom:10px;background:#fff}';
    SET @Css += N'details[open]{padding-bottom:14px}';
    SET @Css += N'summary{cursor:pointer;list-style:none;padding:12px 16px;display:flex;align-items:center;gap:12px;font-weight:700;color:var(--ink2)}';
    SET @Css += N'summary::-webkit-details-marker{display:none}';
    SET @Css += N'summary::before{content:"+";width:18px;height:18px;border-radius:50%;background:var(--soft);color:var(--red);display:inline-flex;align-items:center;justify-content:center;font-size:14px;flex:none}';
    SET @Css += N'details[open] summary::before{content:"\2212"}';
    SET @Css += N'summary .g{color:var(--red);font-size:12px;letter-spacing:.05em}';
    SET @Css += N'summary .c{margin-left:auto;font-weight:400;font-size:12px;color:var(--muted)}';
    SET @Css += N'details .desc,details .tw,details .note{margin:0 16px}';
    SET @Css += N'details .desc{color:var(--muted);font-size:13px;margin-bottom:10px}';
    SET @Css += N'.note{background:var(--soft);border-left:3px solid var(--info);padding:10px 14px;border-radius:4px;font-size:13px}';
    SET @Css += N'.note.warn{border-left-color:var(--warn)}';
    SET @Css += N'.foot{margin-top:48px;padding-top:18px}';
    SET @Css += N'.foot img{display:block;width:100%;max-width:1106px;height:auto}';
    SET @Css += N'.gen{margin-top:10px;font-size:11px;color:var(--muted);text-align:right}';
    SET @Css += N'@media (max-width:900px){.page{padding:36px 20px 20px 34px;background-size:120px auto}.page::before{width:18px;height:110px;top:60px}.cover{padding-right:0}.meta,.counts,.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.exec,.grid2{grid-template-columns:1fr}.cover h1{font-size:24px}}';
    SET @Css += N'@media print{html,body{background:#fff}body{padding:0;font-size:11px}.page{box-shadow:none;max-width:none;padding:10mm 8mm 10mm 14mm}.toc,.chips{display:none}details{break-inside:auto}details summary::before{display:none}section{break-inside:auto}h2{break-after:avoid}.card,.kpi,.cnt,tr{break-inside:avoid}table.data td{white-space:normal;max-width:none}.tw{overflow:visible;border:0}}';
    SET @Css += N'@page{size:A4;margin:10mm}';
    SET @Css += N'.kv tbody th{vertical-align:top;padding:7px 10px;background:#FAFAF8;color:var(--ink2);font-weight:700;font-size:12px;letter-spacing:0;border-top:1px solid var(--line);width:1%;white-space:nowrap}';
    SET @Css += N'table.kv td{white-space:normal}';
    SET @Css += N'.sub{margin:-6px 0 10px;font-size:12px;color:var(--muted)}';
    SET @Css += N'.cpu svg{height:auto}';
    SET @Css += N'p.empty{margin:0}';
    SET @Css += N'details .note{margin-top:10px}';
    SET @Css += N'.tools{display:flex;gap:8px;margin:-4px 0 12px}';
    SET @Css += N'.tools button{font:inherit;font-size:12px;border:1px solid var(--line);background:#fff;color:var(--ink2);border-radius:6px;padding:4px 12px;cursor:pointer}';
    SET @Css += N'.tools button:hover{border-color:var(--red);color:var(--red)}';
    SET @Css += N'@media print{.tools{display:none}}';
    /* V11: guias */
    SET @Css += N'.tabs{position:sticky;top:0;z-index:5;display:flex;flex-wrap:wrap;gap:0 2px;background:#fff;border-bottom:1px solid var(--line);margin:0 -8px;padding:0 8px}';
    SET @Css += N'.tabs button{font:inherit;font-size:12.5px;color:var(--muted);background:none;border:0;border-bottom:3px solid transparent;padding:12px 7px 10px;cursor:pointer;white-space:nowrap}';
    SET @Css += N'.tabs button:hover{color:var(--ink2)}';
    SET @Css += N'.tabs button.on{color:var(--red);border-bottom-color:var(--red);font-weight:700}';
    SET @Css += N'.tabs button .n{display:inline-block;min-width:17px;height:17px;line-height:17px;border-radius:9px;font-size:10px;font-weight:700;text-align:center;margin-left:4px;padding:0 4px}';
    SET @Css += N'.tabs button .n.crit{background:var(--crit-bg);color:var(--crit)}.tabs button .n.warn{background:var(--warn-bg);color:var(--warn)}';
    SET @Css += N'.tabpane{display:none}.tabpane.on{display:block}';
    SET @Css += N'.tabpane section:first-child{margin-top:28px}';
    SET @Css += N'.sub-h{font-size:15px;font-weight:700;color:var(--ink2);margin:26px 0 10px}';
    SET @Css += N'@media print{.tabs{display:none}.tabpane{display:block!important;break-before:page}.tabpane:first-of-type{break-before:auto}}';
    SET @ImgFooter += N'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';
    SET @ImgFooter += N'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';
    SET @ImgFooter += N'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';
    SET @ImgFooter += N'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';
    SET @ImgFooter += N'+prtWsLapjUOmiJU6fvyvdWfpkM06biEpyX1QddnJjq+QhTKtZ0wb7yT3Yk6jvt4jCjvz0JEjZg27ck5zTGzNyXJLiGEEEIImcLq9WUDC9h556Xg+14DQKvXC0a+79UR5VcdGZoBjGPr1tu7kJ2EfN+rIUpkq/++kQS1NUSb2dzIeW7kPHmpwzze6PWChu97K4iiZrp2XhYVhmHqzonnSDduVx8RSOoiknwN4E9pkMm3iJS3LqYkjhUh5YMWcTi0yGMjXjNDa9cTQgghhBBCCCFz4PteCOC3Xi+o+b7XFb2gi8hRoyXFRrgTUury+ghR9Eyr1wtWfN+7QCSk5OXvlrzelr9riISRupxH17kD4FAElBDRl9NdOe+67D50y5cz9u9WRPG8SssWOCTcZwfAjudVVmBtXyxcTMunout9op4ohBBCCCGEEEIIWTzf+763bh0b9XpBwzrWlXIt8TZZAfBGtjL+Brjd1vgbRGJIDUC71wt2fN9rA/gdd54uuo41TOoXXfl9T0QBZhRSTBEFsq2x7LRTc4gqN8bJpyL1ae+UFVNM4VgiT4UwDN/TCoQQQgghhBCyFH5BtPPvjXHsW/EQAYCfHaIKJBznT0ShUn/J+xsA/u71ggsJ6xnFnPN33/cA4G9MCil53I+uueWLOTrZlt81SIyS51VqsySFNcquSz1mvYQQQgghhBBCCHkdNBAJGN8axz71eoGSn8aU93YBfI+7kCD9t2Yl5n3fAfhB/r4xjq/J7x3XmzIJKeIt0pJ/a7hLklPDnahym3wl7m9EOVC0wWpSd4ueKIQQQgghhBBCyOtBQmga1uG873sN+alPeXtXfl8Yf5vH6pLQ9hDAb9Z5W/JnzTg2QiSi/OT7Xt4+2RfzdjYIzlZiRJVZ6mgFwdkKhw4hhBBCCCGEEPLq+AQAvV5wiEjouEDkITJC5KyxjsnQG/26pit1dI2/21JnG5Eosi6v7TjqaBj1fzLe9xH3dwmabdeeWfC8ShgEZ2ra34QQQgghhBBCCCHPiaUJKYQQQgghhBBCCCEvjS9oAkIIIYQQQgghhJB0/H/uu45x2Xk9OwAAAABJRU5ErkJggg==';

    /* ------------------------------------------------------------------------
       Achados consolidados (checklist completo ou, no modo 2, o dashboard)
    ------------------------------------------------------------------------ */
    CREATE TABLE #Achados
    (
        Ordem int NOT NULL,
        Categoria nvarchar(80) NOT NULL,
        Indicador nvarchar(150) NOT NULL,
        Valor nvarchar(500) NULL,
        Status nvarchar(20) NOT NULL,
        Observacao nvarchar(1000) NULL,
        Guia varchar(20) NULL,          /* V11: guia do relatório */
        Linha nvarchar(max) NULL        /* V11: linha HTML pronta */
    );

    IF OBJECT_ID('tempdb..#AssessmentSummary') IS NOT NULL
        INSERT INTO #Achados (Ordem, Categoria, Indicador, Valor, Status, Observacao)
        SELECT ID, Categoria, Indicador, Valor, Status, Observacao FROM #AssessmentSummary;
    ELSE IF OBJECT_ID('tempdb..#Dashboard') IS NOT NULL
        INSERT INTO #Achados (Ordem, Categoria, Indicador, Valor, Status, Observacao)
        SELECT CONVERT(int, ID), Categoria, Indicador, Valor, Status, Observacao FROM #Dashboard WHERE ID NOT IN ('26','27');

    DECLARE @nCrit int, @nWarn int, @nOk int, @nInfo int, @Pct int, @nErros int;
    SELECT @nCrit = ISNULL(SUM(CASE WHEN Status = N'CRÍTICO' THEN 1 ELSE 0 END), 0),
           @nWarn = ISNULL(SUM(CASE WHEN Status = N'ATENÇÃO' THEN 1 ELSE 0 END), 0),
           @nOk   = ISNULL(SUM(CASE WHEN Status = N'OK' THEN 1 ELSE 0 END), 0),
           @nInfo = ISNULL(SUM(CASE WHEN Status NOT IN (N'CRÍTICO', N'ATENÇÃO', N'OK') THEN 1 ELSE 0 END), 0)
    FROM #Achados;
    SET @Pct = CASE WHEN @nOk + @nWarn + @nCrit = 0 THEN 100
                    ELSE CONVERT(int, ROUND(100.0 * @nOk / (@nOk + @nWarn + @nCrit), 0)) END;
    SET @nErros = (SELECT COUNT(*) FROM #ColetaErros);

    DECLARE @TemDetalhe bit = CASE WHEN EXISTS (SELECT 1 FROM #Secoes) THEN 1 ELSE 0 END;
    DECLARE @DataColeta nvarchar(20) = CONVERT(nvarchar(10), @AssessmentDate, 103) + N' ' + LEFT(CONVERT(nvarchar(8), @AssessmentDate, 108), 5);
    DECLARE @Collation nvarchar(128) = CONVERT(nvarchar(128), SERVERPROPERTY('Collation'));
    DECLARE @UpdateLevel nvarchar(128) = CONVERT(nvarchar(128), SERVERPROPERTY('ProductUpdateLevel'));
    DECLARE @Autenticacao nvarchar(40) = CASE WHEN SERVERPROPERTY('IsIntegratedSecurityOnly') = 1 THEN N'Somente Windows' ELSE N'Windows + SQL Server (misto)' END;
    DECLARE @Sec int = 0;

    /* ------------------------------------------------------------------------
       V11: organização em guias (abas)
       Cada linha do checklist e cada seção de detalhe é atribuída a uma guia.
    ------------------------------------------------------------------------ */
    UPDATE #Achados SET
        Guia = CASE WHEN Indicador = N'xp_cmdshell' THEN 'seguranca'
                    ELSE CASE LEFT(Categoria, 2)
                        WHEN '01' THEN 'geral'         WHEN '02' THEN 'geral'
                        WHEN '03' THEN 'ha'            WHEN '04' THEN 'ha'            WHEN '23' THEN 'ha'
                        WHEN '05' THEN 'conexoes'      WHEN '06' THEN 'conexoes'      WHEN '21' THEN 'conexoes'
                        WHEN '07' THEN 'databases'     WHEN '09' THEN 'databases'     WHEN '17' THEN 'databases'
                        WHEN '08' THEN 'armazenamento'
                        WHEN '10' THEN 'backups'       WHEN '16' THEN 'backups'       WHEN '22' THEN 'backups'
                        WHEN '11' THEN 'jobs'
                        WHEN '12' THEN 'seguranca'     WHEN '18' THEN 'seguranca'
                        WHEN '13' THEN 'performance'   WHEN '14' THEN 'performance'   WHEN '15' THEN 'performance'
                        WHEN '19' THEN 'performance'   WHEN '20' THEN 'performance'
                        ELSE 'geral' END END,
        Linha = N'<tr><td>' + REPLACE(REPLACE(REPLACE(Indicador, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td>'
              + N'<td class="n">' + ISNULL(NULLIF(REPLACE(REPLACE(REPLACE(ISNULL(Valor, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N''), N'—') + N'</td>'
              + N'<td><span class="b ' + CASE Status WHEN N'OK' THEN N'ok' WHEN N'ATENÇÃO' THEN N'warn' WHEN N'CRÍTICO' THEN N'crit' ELSE N'info' END + N'">' + Status + N'</span></td>'
              + N'<td class="obs">' + ISNULL(NULLIF(REPLACE(REPLACE(REPLACE(ISNULL(Observacao, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N''), N'—') + N'</td></tr>';

    UPDATE #V11_HtmlDet SET Guia =
        CASE WHEN Titulo LIKE N'%Buffer pool%' OR Titulo LIKE N'%Trace flag%' THEN 'performance'
             WHEN Titulo LIKE N'%log de transações%' OR Titulo LIKE N'%VLF%' THEN 'databases'
             WHEN Grupo = N'20' AND Titulo LIKE N'Objetos%' THEN 'databases'
             ELSE CASE Grupo
                WHEN N'01' THEN 'geral'         WHEN N'02' THEN 'geral'
                WHEN N'03' THEN 'ha'            WHEN N'04' THEN 'ha'            WHEN N'25' THEN 'ha'
                WHEN N'05' THEN 'conexoes'      WHEN N'06' THEN 'conexoes'      WHEN N'22' THEN 'conexoes'
                WHEN N'07' THEN 'databases'     WHEN N'09' THEN 'databases'     WHEN N'17' THEN 'databases'
                WHEN N'08' THEN 'armazenamento'
                WHEN N'10' THEN 'backups'       WHEN N'16' THEN 'backups'       WHEN N'23' THEN 'backups'
                WHEN N'11' THEN 'jobs'          WHEN N'24' THEN 'jobs'
                WHEN N'12' THEN 'seguranca'     WHEN N'19' THEN 'seguranca'
                WHEN N'13' THEN 'performance'   WHEN N'14' THEN 'performance'   WHEN N'15' THEN 'performance'
                WHEN N'18' THEN 'performance'   WHEN N'20' THEN 'performance'   WHEN N'21' THEN 'performance'
                ELSE 'geral' END END;

    DECLARE @Guias TABLE (Pos int PRIMARY KEY, Id varchar(20) NOT NULL, Rotulo nvarchar(60) NOT NULL, Curto nvarchar(40) NOT NULL, Lead nvarchar(400) NULL);
    INSERT INTO @Guias (Pos, Id, Rotulo, Curto, Lead) VALUES
        (1,  'geral',         N'Visão Geral',           N'Visão Geral',          NULL),
        (2,  'atencao',       N'Pontos de Atenção',     N'Atenção',              NULL),
        (3,  'databases',     N'Databases',             N'Databases',            N'Configuração, tamanho e estado das databases, objetos, TempDB e arquivos de log (VLFs).'),
        (4,  'performance',   N'Performance',           N'Performance',          N'Uso de CPU, memória, I/O, buffer pool, esperas, índices, queries mais pesadas e configurações que afetam o desempenho.'),
        (5,  'backups',       N'Backups e Integridade', N'Backups',              N'Histórico de backups, último CHECKDB e páginas suspeitas.'),
        (6,  'seguranca',     N'Segurança',             N'Segurança',            N'Logins, membros de server roles, permissões, política de senha e configurações de segurança.'),
        (7,  'jobs',          N'Jobs e Alertas',        N'Jobs',                 N'Jobs do SQL Server Agent, agendamentos, resultado da última execução e alertas configurados.'),
        (8,  'armazenamento', N'Armazenamento',         N'Armazenamento',        N'Volumes, maiores databases e localização dos arquivos de dados e log.'),
        (9,  'conexoes',      N'Conexões',              N'Conexões',             N'Linked servers, endpoints e sessões de usuário no momento da coleta.'),
        (10, 'ha',            N'Alta Disponibilidade',  N'Alta Disponibilidade', N'Cluster (FCI), Always On Availability Groups e Log Shipping.'),
        (11, 'checklist',     N'Checklist',             N'Checklist',            NULL);

    DECLARE @Nav nvarchar(max), @CapGrid nvarchar(max) = N'', @CapCpu nvarchar(max) = N'', @Pane nvarchar(max),
            @gPos int, @gId varchar(20), @gRotulo nvarchar(60), @gLead nvarchar(400);

    SELECT @Nav = STRING_AGG(CAST(
                N'<button type="button" data-t="' + g.Id + N'" onclick="tab(''' + g.Id + N''')">' + g.Curto
              + CASE WHEN g.Id <> 'checklist' AND n.c + n.w > 0
                     THEN N'<span class="n ' + CASE WHEN n.c > 0 THEN N'crit' ELSE N'warn' END + N'">' + CONVERT(nvarchar(10), n.c + n.w) + N'</span>'
                     ELSE N'' END
              + N'</button>' AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY g.Pos)
    FROM @Guias AS g
    CROSS APPLY (SELECT ISNULL(SUM(CASE WHEN a.Status = N'CRÍTICO' THEN 1 ELSE 0 END), 0) AS c,
                        ISNULL(SUM(CASE WHEN a.Status = N'ATENÇÃO' THEN 1 ELSE 0 END), 0) AS w
                 FROM #Achados AS a
                 WHERE a.Guia = g.Id OR g.Id = 'atencao') AS n;

    /* ------------------------------------------------------------------------
       Cabeçalho / capa
    ------------------------------------------------------------------------ */
    SET @Html  = N'<html lang="pt-BR"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>';
    SET @Html += N'<title>Assessment SQL Server - ' + REPLACE(REPLACE(REPLACE(ISNULL(@ServerName, N'N/D'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</title>';
    SET @Html += N'<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Lato:wght@400;700&amp;display=swap"/>';
    SET @Html += N'<style>' + @Css + N'</style>';
    SET @Html += N'<script>function tg(o){document.querySelectorAll("details").forEach(function(d){d.open=o})}'
               + N'function ab(){var e=document.getElementById(location.hash.substring(1));if(e){if(e.tagName=="DETAILS"){e.open=true}}}'
               + N'window.addEventListener("hashchange",ab);window.addEventListener("load",ab);window.addEventListener("beforeprint",function(){tg(true)});</script>';
    SET @Html += N'<script>function tab(id){document.querySelectorAll(''.tabpane'').forEach(function(p){p.classList.toggle(''on'',p.id==''t-''+id)});document.querySelectorAll(''.tabs button'').forEach(function(b){b.classList.toggle(''on'',b.getAttribute(''data-t'')==id)});if(history.replaceState){history.replaceState(null,'''',''#''+id)}window.scrollTo(0,document.querySelector(''.tabs'').offsetTop)}'
               + N'window.addEventListener(''load'',function(){var h=location.hash.substring(1);tab(document.getElementById(''t-''+h)?h:''geral'')});window.addEventListener(''beforeprint'',function(){document.querySelectorAll(''details'').forEach(function(d){d.open=true})});</script>';
    SET @Html += N'</head><body><div class="page">';

    SET @Html += N'<header class="cover"><div class="kicker">Assessment técnico · SQL Server</div>'
               + N'<h1>Levantamento do Ambiente de Banco de Dados</h1>'
               + CASE WHEN LTRIM(RTRIM(@Cliente)) <> N'' THEN N'<div class="client">Cliente: <b>' + REPLACE(REPLACE(REPLACE(@Cliente, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>' ELSE N'' END
               + N'<div class="server"><span class="lbl">Servidor \ Instância</span><span class="val">' + REPLACE(REPLACE(REPLACE(ISNULL(@ServerName, N'N/D'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</span></div>';

    SET @Html += N'<div class="meta">'
               + N'<div><span>Data da coleta</span><b>' + @DataColeta + N'</b></div>'
               + N'<div><span>Versão</span><b>' + REPLACE(REPLACE(REPLACE(@Dashboard_SQLServerVersion + N' · ' + ISNULL(@ProductVersion, N'') + N' · ' + ISNULL(@ProductLevel, N'') + ISNULL(N' ' + @UpdateLevel, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>'
               + N'<div><span>Edição</span><b>' + REPLACE(REPLACE(REPLACE(ISNULL(@Edition, N'N/D'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>'
               + N'<div><span>Executado por</span><b>' + REPLACE(REPLACE(REPLACE(ISNULL(SUSER_SNAME(), N'N/D'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>'
               + N'<div><span>Sistema operacional</span><b>' + REPLACE(REPLACE(REPLACE(ISNULL(NULLIF(@Dashboard_HostDistribution, N''), N'N/D') + ISNULL(N' (' + NULLIF(@Dashboard_HostRelease, N'N/D') + N')', N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>'
               + N'<div><span>Virtualização</span><b>' + REPLACE(REPLACE(REPLACE(ISNULL(@Dashboard_VirtualMachine, N'NONE'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>'
               + N'<div><span>Collation</span><b>' + REPLACE(REPLACE(REPLACE(ISNULL(@Collation, N'N/D'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</b></div>'
               + N'<div><span>Autenticação</span><b>' + @Autenticacao + N'</b></div>'
               + N'</div></header>';

    SET @Html += N'<nav class="tabs">' + ISNULL(@Nav, N'') + N'</nav>';

    /* ------------------------------------------------------------------------
       01 - Resumo executivo
    ------------------------------------------------------------------------ */
    SET @Html += N'<div class="tabpane" id="t-geral"><section id="resumo"><h2><span class="num">01</span>Resumo executivo</h2><div class="exec"><div class="gauge">'
               + N'<svg viewBox="0 0 140 140"><circle cx="70" cy="70" r="52" fill="none" stroke="#E4E4E1" stroke-width="14"></circle>'
               + N'<circle cx="70" cy="70" r="52" fill="none" stroke="'
               + CASE WHEN @Pct >= 80 THEN N'#2E7D4F' WHEN @Pct >= 50 THEN N'#A86400' ELSE N'#B3121B' END
               + N'" stroke-width="14" stroke-linecap="round" stroke-dasharray="'
               + CONVERT(nvarchar(20), CONVERT(decimal(7,1), 326.73 * @Pct / 100.0)) + N' 326.73" transform="rotate(-90 70 70)"></circle>'
               + N'<text x="70" y="80" text-anchor="middle" class="pct">' + CONVERT(nvarchar(10), @Pct) + N'%</text></svg>'
               + N'<div class="cap">das verificações automáticas em conformidade</div></div><div class="counts">'
               + N'<div class="cnt crit"><span>Crítico</span><b>' + CONVERT(nvarchar(10), @nCrit) + N'</b><em>ação imediata</em></div>'
               + N'<div class="cnt warn"><span>Atenção</span><b>' + CONVERT(nvarchar(10), @nWarn) + N'</b><em>revisar</em></div>'
               + N'<div class="cnt ok"><span>OK</span><b>' + CONVERT(nvarchar(10), @nOk) + N'</b><em>em conformidade</em></div>'
               + N'<div class="cnt"><span>Info</span><b>' + CONVERT(nvarchar(10), @nInfo) + N'</b><em>informativo</em></div></div></div>';

    SET @Html += N'<div class="kpis">'
               + N'<div class="kpi"><span class="k">CPUs lógicas</span><span class="v">' + ISNULL(CONVERT(nvarchar(10), @Dashboard_LogicalCPU), N'-') + N'</span><span class="s">'
               + ISNULL(CONVERT(nvarchar(10), @K_Sockets), N'?') + N' socket(s) · ' + ISNULL(CONVERT(nvarchar(10), @K_CoresPorSocket), N'?') + N' cores/socket · '
               + ISNULL(CONVERT(nvarchar(10), @K_Numa), N'?') + N' NUMA</span></div>'
               + N'<div class="kpi"><span class="k">Memória física</span><span class="v">' + ISNULL(FORMAT(@Dashboard_PhysicalMemoryMB / 1024.0, N'N1', N'pt-BR'), N'-') + N' GB</span><span class="s">'
               + CASE WHEN @K_MaxServerMemoryMB >= 2147483647 THEN N'max server memory: não configurado'
                      ELSE N'max server memory: ' + ISNULL(FORMAT(@K_MaxServerMemoryMB / 1024.0, N'N1', N'pt-BR'), N'?') + N' GB' END + N'</span></div>'
               + N'<div class="kpi"><span class="k">Databases de usuário</span><span class="v">' + ISNULL(CONVERT(nvarchar(10), @Dashboard_Databases), N'-') + N'</span><span class="s">'
               + CASE WHEN @Dashboard_DatabasesNotOnline > 0 THEN CONVERT(nvarchar(10), @Dashboard_DatabasesNotOnline) + N' fora de ONLINE' ELSE N'todas ONLINE' END + N'</span></div>'
               + N'<div class="kpi"><span class="k">Tamanho alocado</span><span class="v">' + ISNULL(FORMAT(@K_DataGB + @K_LogGB, N'N1', N'pt-BR'), N'-') + N' GB</span><span class="s">dados '
               + ISNULL(FORMAT(@K_DataGB, N'N1', N'pt-BR'), N'-') + N' GB · log ' + ISNULL(FORMAT(@K_LogGB, N'N1', N'pt-BR'), N'-') + N' GB</span></div>';

    SET @Html += N'<div class="kpi"><span class="k">Always On</span><span class="v">' + CASE WHEN @Dashboard_Hadr = 1 THEN N'Habilitado' ELSE N'Não habilitado' END + N'</span><span class="s">'
               + CONVERT(nvarchar(10), @Dashboard_AGs) + N' AG(s) · ' + CONVERT(nvarchar(10), @Dashboard_Replicas) + N' réplica(s)'
               + CASE WHEN @IsClustered = 1 THEN N' · instância FCI' ELSE N'' END
               + CASE WHEN @K_LSPrim + @K_LSSec > 0 THEN N' · log shipping: ' + CONVERT(nvarchar(10), @K_LSPrim) + N' primária(s), ' + CONVERT(nvarchar(10), @K_LSSec) + N' secundária(s)'
                      ELSE N' · sem log shipping' END + N'</span></div>'
               + N'<div class="kpi"><span class="k">Jobs do SQL Agent</span><span class="v">' + ISNULL(CONVERT(nvarchar(10), @Dashboard_Jobs), N'-') + N'</span><span class="s">'
               + CASE WHEN @K_JobsFalha > 0 THEN CONVERT(nvarchar(10), @K_JobsFalha) + N' com falha na última execução' ELSE N'nenhum com falha na última execução' END + N'</span></div>'
               + N'<div class="kpi"><span class="k">Linked servers</span><span class="v">' + ISNULL(CONVERT(nvarchar(10), @Dashboard_LinkedServers), N'-') + N'</span><span class="s">não sistêmicos</span></div>'
               + N'<div class="kpi"><span class="k">Uptime</span><span class="v">' + ISNULL(CONVERT(nvarchar(10), DATEDIFF(day, @K_StartTime, GETDATE())), N'-') + N' dias</span><span class="s">desde '
               + ISNULL(CONVERT(nvarchar(10), @K_StartTime, 103) + N' ' + LEFT(CONVERT(nvarchar(8), @K_StartTime, 108), 5), N'-') + N'</span></div>'
               + N'</div></section>';

    SELECT @Tmp = NULL, @Tmp2 = NULL;
    SELECT @Tmp = STRING_AGG(CAST(a.Linha AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY a.Ordem)
    FROM #Achados AS a WHERE a.Guia = 'geral';
    SELECT @Tmp2 = STRING_AGG(CAST(d.Html AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY d.Ordem)
    FROM #V11_HtmlDet AS d WHERE d.Guia = 'geral';
    IF @Tmp IS NOT NULL OR @Tmp2 IS NOT NULL
        SET @Html += N'<section>'
                   + CASE WHEN @Tmp IS NULL THEN N''
                          ELSE N'<div class="sub-h">Verificações desta área</div><div class="tw"><table><thead><tr><th>Indicador</th><th>Valor</th><th>Status</th><th>Observação</th></tr></thead><tbody>'
                               + @Tmp + N'</tbody></table></div>' END
                   + CASE WHEN @Tmp2 IS NULL THEN N''
                          ELSE N'<div class="sub-h">Detalhamento técnico</div>'
                               + N'<div class="tools"><button type="button" onclick="tg(true)">Expandir tudo</button><button type="button" onclick="tg(false)">Recolher tudo</button></div>'
                               + @Tmp2 END
                   + N'</section>';
    SET @Html += N'</div>';

    /* ------------------------------------------------------------------------
       02 - Pontos de atenção
    ------------------------------------------------------------------------ */
    SET @Tmp = NULL;
    SELECT @Tmp = STRING_AGG(CAST(
                N'<tr><td><span class="b ' + CASE a.Status WHEN N'CRÍTICO' THEN N'crit' ELSE N'warn' END + N'">' + a.Status + N'</span></td>'
              + N'<td>' + REPLACE(REPLACE(REPLACE(a.Categoria, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td><td>' + REPLACE(REPLACE(REPLACE(a.Indicador, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td>'
              + N'<td class="n">' + ISNULL(NULLIF(REPLACE(REPLACE(REPLACE(ISNULL(a.Valor, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N''), N'—') + N'</td>'
              + N'<td class="obs">' + ISNULL(NULLIF(REPLACE(REPLACE(REPLACE(ISNULL(a.Observacao, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N''), N'—') + N'</td></tr>'
           AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY CASE a.Status WHEN N'CRÍTICO' THEN 0 ELSE 1 END, a.Ordem)
    FROM #Achados AS a
    WHERE a.Status IN (N'CRÍTICO', N'ATENÇÃO');

    SET @Html += N'<div class="tabpane" id="t-atencao"><section id="atencao"><h2><span class="num">02</span>Pontos de atenção</h2>'
               + N'<p class="lead">Itens classificados automaticamente como CRÍTICO ou ATENÇÃO. A classificação é indicativa e deve ser validada com o contexto do ambiente e do negócio.</p>'
               + CASE WHEN @Tmp IS NULL
                      THEN N'<div class="note">Nenhum ponto de atenção identificado pelas verificações automáticas.</div>'
                      ELSE N'<div class="tw"><table><thead><tr><th>Status</th><th>Categoria</th><th>Indicador</th><th>Valor</th><th>Observação</th></tr></thead><tbody>' + @Tmp + N'</tbody></table></div>' END
               + N'</section></div>';

    /* ------------------------------------------------------------------------
       03 - Capacidade (volumes, maiores databases, CPU)
    ------------------------------------------------------------------------ */
    SET @Tmp = NULL;
    BEGIN TRY
        SELECT @Tmp = STRING_AGG(CAST(
                    N'<span class="n">' + REPLACE(REPLACE(REPLACE(v.mp, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</span><div class="bar"><i class="' + k.cls + N'" style="width:' + CONVERT(nvarchar(10), m.usedpct) + N'%">'
                  + CONVERT(nvarchar(10), m.usedpct) + N'</i></div><span class="t">' + FORMAT(m.usedGB, N'N1', N'pt-BR') + N' / ' + FORMAT(m.totalGB, N'N1', N'pt-BR')
                  + N' GB · ' + CONVERT(nvarchar(10), m.freepct) + N'% livre</span>'
               AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY v.mp)
        FROM
        (
            SELECT DISTINCT vs.volume_mount_point AS mp, vs.total_bytes, vs.available_bytes
            FROM sys.master_files AS mf
            JOIN sys.databases AS d ON d.database_id = mf.database_id AND d.state_desc = 'ONLINE'
            CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) AS vs
        ) AS v
        CROSS APPLY (SELECT CONVERT(int, ROUND(100.0 - v.available_bytes * 100.0 / NULLIF(v.total_bytes, 0), 0)) AS usedpct,
                            CONVERT(int, ROUND(v.available_bytes * 100.0 / NULLIF(v.total_bytes, 0), 0)) AS freepct,
                            (v.total_bytes - v.available_bytes) / 1073741824.0 AS usedGB,
                            v.total_bytes / 1073741824.0 AS totalGB) AS m
        CROSS APPLY (SELECT CASE WHEN m.freepct < 10 THEN N'crit' WHEN m.freepct < 15 THEN N'warn' ELSE N'ok' END AS cls) AS k;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'Capacidade - volumes (HTML)', ERROR_MESSAGE());
    END CATCH;

    SET @Tmp2 = NULL;
    BEGIN TRY
        ;WITH s AS
        (
            SELECT TOP (10) database_id, SUM(CONVERT(bigint, size)) / 128.0 / 1024.0 AS gb
            FROM sys.master_files
            GROUP BY database_id
            ORDER BY SUM(CONVERT(bigint, size)) DESC
        ), mx AS (SELECT MAX(gb) AS mx FROM s)
        SELECT @Tmp2 = STRING_AGG(CAST(
                    N'<span class="n">' + REPLACE(REPLACE(REPLACE(ISNULL(DB_NAME(s.database_id), N'database_id ' + CONVERT(nvarchar(10), s.database_id)), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</span>'
                  + N'<div class="bar"><i class="db" style="width:' + CONVERT(nvarchar(10), w.pct) + N'%">' + CONVERT(nvarchar(10), w.pct) + N'</i></div>'
                  + N'<span class="t">' + FORMAT(s.gb, CASE WHEN s.gb < 10 THEN N'N2' ELSE N'N1' END, N'pt-BR') + N' GB</span>'
               AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY s.gb DESC)
        FROM s
        CROSS JOIN mx
        CROSS APPLY (SELECT CASE WHEN s.gb * 100.0 / NULLIF(mx.mx, 0) < 1 THEN 1
                                 ELSE CONVERT(int, ROUND(s.gb * 100.0 / NULLIF(mx.mx, 0), 0)) END AS pct) AS w;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'Capacidade - databases (HTML)', ERROR_MESSAGE());
    END CATCH;

    SET @CapGrid = N'<div class="grid2">'
               + N'<div class="card"><h3>Volumes com arquivos do SQL Server</h3>'
               + CASE WHEN @Tmp IS NULL THEN N'<p class="empty">Informação de volumes indisponível.</p>' ELSE N'<div class="hb">' + @Tmp + N'</div>' END
               + N'</div><div class="card"><h3>Maiores databases (dados + log)</h3>'
               + CASE WHEN @Tmp2 IS NULL THEN N'<p class="empty">Informação indisponível.</p>' ELSE N'<div class="hb">' + @Tmp2 + N'</div>' END
               + N'</div></div>';

    /* Histórico de CPU do ring buffer (1 amostra por minuto, até 256 min) */
    DECLARE @PtsSql nvarchar(max), @PtsOth nvarchar(max), @CpuN int, @CpuIni datetime, @CpuFim datetime, @CpuMed int, @CpuMax int;
    BEGIN TRY
        ;WITH rb AS
        (
            SELECT TOP (256)
                   y.record_id,
                   y.SQLProcessUtilization AS SqlCpu,
                   CASE WHEN 100 - y.SystemIdle - y.SQLProcessUtilization < 0 THEN 0 ELSE 100 - y.SystemIdle - y.SQLProcessUtilization END AS OthCpu,
                   DATEADD(ms, -1 * (ts.ms_ticks - y.[timestamp]), GETDATE()) AS EventTime
            FROM
            (
                SELECT record.value('(./Record/@id)[1]','int') AS record_id,
                       record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int') AS SystemIdle,
                       record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS SQLProcessUtilization,
                       [timestamp]
                FROM
                (
                    SELECT [timestamp], CONVERT(xml, record) AS record
                    FROM sys.dm_os_ring_buffers
                    WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
                      AND record LIKE N'%<SystemHealth>%'
                ) AS x
            ) AS y
            CROSS JOIN (SELECT ms_ticks FROM sys.dm_os_sys_info) AS ts
            ORDER BY y.record_id DESC
        ), o AS
        (
            SELECT ROW_NUMBER() OVER (ORDER BY record_id) AS i, COUNT(*) OVER () AS n, SqlCpu, OthCpu, EventTime
            FROM rb
        )
        SELECT
            @PtsSql = STRING_AGG(CAST(p.x + N',' + CONVERT(nvarchar(20), CONVERT(decimal(7,1), 10 + 140.0 * (1 - ISNULL(o.SqlCpu, 0) / 100.0))) AS nvarchar(max)), N' ') WITHIN GROUP (ORDER BY o.i),
            @PtsOth = STRING_AGG(CAST(p.x + N',' + CONVERT(nvarchar(20), CONVERT(decimal(7,1), 10 + 140.0 * (1 - ISNULL(o.OthCpu, 0) / 100.0))) AS nvarchar(max)), N' ') WITHIN GROUP (ORDER BY o.i),
            @CpuN = MAX(o.n), @CpuIni = MIN(o.EventTime), @CpuFim = MAX(o.EventTime),
            @CpuMed = AVG(o.SqlCpu), @CpuMax = MAX(o.SqlCpu)
        FROM o
        CROSS APPLY (SELECT CONVERT(nvarchar(20), CONVERT(decimal(7,1),
                        CASE WHEN o.n <= 1 THEN 40 ELSE 40 + (o.i - 1) * 960.0 / (o.n - 1) END)) AS x) AS p;
    END TRY
    BEGIN CATCH
        INSERT INTO #ColetaErros (Secao, Mensagem) VALUES (N'Capacidade - CPU (HTML)', ERROR_MESSAGE());
    END CATCH;

    IF @CpuN > 1
        SET @CapCpu = N'<div class="card cpu"><h3>Uso de CPU · últimos ' + CONVERT(nvarchar(10), @CpuN) + N' minutos</h3>'
                   + N'<p class="sub">SQL Server: média ' + CONVERT(nvarchar(10), @CpuMed) + N'% · pico ' + CONVERT(nvarchar(10), @CpuMax) + N'% (amostras por minuto do ring buffer do scheduler monitor)</p>'
                   + N'<svg viewBox="0 0 1000 180">'
                   + N'<line class="grid" x1="40" x2="1000" y1="10" y2="10"></line><text class="ax" x="32" y="14" text-anchor="end">100%</text>'
                   + N'<line class="grid" x1="40" x2="1000" y1="45" y2="45"></line><text class="ax" x="32" y="49" text-anchor="end">75%</text>'
                   + N'<line class="grid" x1="40" x2="1000" y1="80" y2="80"></line><text class="ax" x="32" y="84" text-anchor="end">50%</text>'
                   + N'<line class="grid" x1="40" x2="1000" y1="115" y2="115"></line><text class="ax" x="32" y="119" text-anchor="end">25%</text>'
                   + N'<line class="grid" x1="40" x2="1000" y1="150" y2="150"></line><text class="ax" x="32" y="154" text-anchor="end">0%</text>'
                   + N'<polyline class="l-oth" points="' + @PtsOth + N'"></polyline>'
                   + N'<polyline class="l-sql" points="' + @PtsSql + N'"></polyline>'
                   + N'<text class="ax" x="40" y="174">' + CONVERT(nvarchar(10), @CpuIni, 103) + N' ' + LEFT(CONVERT(nvarchar(8), @CpuIni, 108), 5) + N'</text>'
                   + N'<text class="ax" x="1000" y="174" text-anchor="end">' + CONVERT(nvarchar(10), @CpuFim, 103) + N' ' + LEFT(CONVERT(nvarchar(8), @CpuFim, 108), 5) + N'</text>'
                   + N'</svg><div class="legend"><span>SQL Server</span><span class="o">Outros processos</span></div></div>';
    ELSE
        SET @CapCpu = N'<div class="card cpu"><h3>Uso de CPU</h3><p class="empty">Histórico de CPU indisponível (ring buffer vazio ou sem permissão).</p></div>';

    /* ------------------------------------------------------------------------
       V11: guias 03 a 10 (verificações da área + capacidade/CPU + detalhamento)
    ------------------------------------------------------------------------ */
    DECLARE curGuias CURSOR LOCAL FAST_FORWARD FOR
        SELECT Pos, Id, Rotulo, Lead FROM @Guias WHERE Pos BETWEEN 3 AND 10 ORDER BY Pos;
    OPEN curGuias;
    FETCH NEXT FROM curGuias INTO @gPos, @gId, @gRotulo, @gLead;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SELECT @Tmp = NULL, @Tmp2 = NULL;

        SELECT @Tmp = STRING_AGG(CAST(a.Linha AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY a.Ordem)
        FROM #Achados AS a WHERE a.Guia = @gId;

        SELECT @Tmp2 = STRING_AGG(CAST(d.Html AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY d.Ordem)
        FROM #V11_HtmlDet AS d WHERE d.Guia = @gId;

        SET @Pane = N'<div class="tabpane" id="t-' + @gId + N'"><section><h2><span class="num">' + RIGHT(N'0' + CONVERT(nvarchar(2), @gPos), 2) + N'</span>' + @gRotulo + N'</h2>'
                  + N'<p class="lead">' + @gLead + N'</p>'
                  + CASE WHEN @Tmp IS NULL THEN N''
                         ELSE N'<div class="sub-h">Verificações desta área</div><div class="tw"><table><thead><tr><th>Indicador</th><th>Valor</th><th>Status</th><th>Observação</th></tr></thead><tbody>'
                              + @Tmp + N'</tbody></table></div>' END
                  + CASE WHEN @gId = 'armazenamento' THEN N'<div class="sub-h">Capacidade</div>' + @CapGrid ELSE N'' END
                  + CASE WHEN @gId = 'performance' THEN N'<div class="sub-h">Uso de CPU</div>' + @CapCpu ELSE N'' END
                  + CASE WHEN @Tmp2 IS NULL THEN N''
                         ELSE N'<div class="sub-h">Detalhamento técnico</div>'
                              + N'<div class="tools"><button type="button" onclick="tg(true)">Expandir tudo</button><button type="button" onclick="tg(false)">Recolher tudo</button></div>'
                              + @Tmp2 END
                  + N'</section></div>';
        SET @Html += @Pane;

        FETCH NEXT FROM curGuias INTO @gPos, @gId, @gRotulo, @gLead;
    END;
    CLOSE curGuias;
    DEALLOCATE curGuias;

    /* ------------------------------------------------------------------------
       11 - Checklist completo
    ------------------------------------------------------------------------ */
    SET @Tmp = NULL;
    SELECT @Tmp = STRING_AGG(CAST(
                CASE WHEN a.rn = 1 THEN N'<tr class="grp"><td colspan="4">' + REPLACE(REPLACE(REPLACE(a.Categoria, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td></tr>' ELSE N'' END
              + N'<tr><td>' + REPLACE(REPLACE(REPLACE(a.Indicador, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td>'
              + N'<td class="n">' + ISNULL(NULLIF(REPLACE(REPLACE(REPLACE(ISNULL(a.Valor, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N''), N'—') + N'</td>'
              + N'<td><span class="b ' + CASE a.Status WHEN N'OK' THEN N'ok' WHEN N'ATENÇÃO' THEN N'warn' WHEN N'CRÍTICO' THEN N'crit' ELSE N'info' END + N'">' + a.Status + N'</span></td>'
              + N'<td class="obs">' + ISNULL(NULLIF(REPLACE(REPLACE(REPLACE(ISNULL(a.Observacao, N''), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;'), N''), N'—') + N'</td></tr>'
           AS nvarchar(max)), N'') WITHIN GROUP (ORDER BY a.CatOrdem, a.Ordem)
    FROM
    (
        SELECT *, ROW_NUMBER() OVER (PARTITION BY Categoria ORDER BY Ordem) AS rn,
                  MIN(Ordem) OVER (PARTITION BY Categoria) AS CatOrdem
        FROM #Achados
    ) AS a;

    SET @Html += N'<div class="tabpane" id="t-checklist"><section id="checklist"><h2><span class="num">11</span>Checklist do assessment</h2>'
               + N'<p class="lead">Todas as verificações executadas, agrupadas por categoria.</p>'
               + N'<div class="tw"><table><thead><tr><th>Indicador</th><th>Valor</th><th>Status</th><th>Observação</th></tr></thead><tbody>'
               + ISNULL(@Tmp, N'<tr><td class="empty" colspan="4">Sem dados.</td></tr>') + N'</tbody></table></div></section>';

    SET @Sec = 11;

    /* ------------------------------------------------------------------------
       Observações da coleta (falhas de permissão/DMV)
    ------------------------------------------------------------------------ */
    SET @nErros = (SELECT COUNT(*) FROM #ColetaErros);
    IF @nErros > 0
    BEGIN
        SET @Sec += 1;
        SET @Tmp = NULL;
        SELECT @Tmp = STRING_AGG(CAST(N'<tr><td>' + REPLACE(REPLACE(REPLACE(Secao, N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td><td class="obs">' + REPLACE(REPLACE(REPLACE(ISNULL(Mensagem, N'-'), N'&', N'&amp;'), N'<', N'&lt;'), N'>', N'&gt;') + N'</td></tr>' AS nvarchar(max)), N'')
                      WITHIN GROUP (ORDER BY ID)
        FROM #ColetaErros;
        SET @Html += N'<section id="coleta"><h2><span class="num">' + CONVERT(nvarchar(2), @Sec) + N'</span>Observações da coleta</h2>'
                   + N'<p class="lead">Itens que não puderam ser coletados (geralmente por permissão da conta de execução). As demais seções não foram afetadas.</p>'
                   + N'<div class="tw"><table><thead><tr><th>Seção</th><th>Mensagem</th></tr></thead><tbody>' + @Tmp + N'</tbody></table></div></section>';
    END;

    SET @Html += N'</div>'; /* fim da guia Checklist */

    /* ------------------------------------------------------------------------
       Rodapé (papel timbrado)
    ------------------------------------------------------------------------ */
    SET @Html += N'<footer class="foot"><img alt="Solo Network - Rod. João Leopoldo Jacomel, 12.162 - Pinhais/PR - 0800 604 9596 - www.solonetwork.com.br" src="data:image/png;base64,' + @ImgFooter + N'"/>'
               + N'<div class="gen">SQL Server Assessment ' + @ScriptVersao + N' · gerado em ' + @DataColeta + N' · coleta somente leitura</div></footer>'
               + N'</div></body></html>';

    /* Remove caracteres de controle inválidos em XML (ex.: mensagens de jobs) */
    DECLARE @cc int = 1;
    WHILE @cc < 32
    BEGIN
        IF @cc NOT IN (9, 10, 13)
            SET @Html = REPLACE(@Html COLLATE Latin1_General_BIN2, NCHAR(@cc), N' ');
        SET @cc += 1;
    END;

    /* ------------------------------------------------------------------------
       SAÍDA: último resultset
       - Coluna XML clicável no SSMS (abrir > Salvar Como .html)
       - Se o conteúdo não for XML válido, retorna o HTML em partes de texto
    ------------------------------------------------------------------------ */
    DECLARE @HtmlXml xml = TRY_CONVERT(xml, @Html, 1);

    IF @HtmlXml IS NOT NULL
        SELECT @HtmlXml AS Relatorio_HTML;
    ELSE
    BEGIN
        PRINT 'AVISO: o HTML não pôde ser convertido para XML. Retornando em partes de texto (coluna Relatorio_HTML_Texto).';
        ;WITH n AS
        (
            SELECT TOP (CONVERT(int, CEILING(DATALENGTH(@Html) / 2 / 4000.0)))
                   ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS Parte
            FROM sys.all_columns AS a CROSS JOIN sys.all_columns AS b
        )
        SELECT Parte, SUBSTRING(@Html, (Parte - 1) * 4000 + 1, 4000) AS Relatorio_HTML_Texto
        FROM n
        ORDER BY Parte;
    END;

    PRINT 'RELATÓRIO HTML: clique na coluna Relatorio_HTML do último resultset e salve como .html';
END;

/* ============================================================================
   FINALIZAÇÃO
============================================================================= */
PRINT '=================================================================';
PRINT 'ASSESSMENT FINALIZADO';
PRINT 'Modo 1 = Dashboard + detalhes | Modo 2 = Dashboard | Modo 3 = detalhes.';
PRINT 'O Dashboard é o primeiro resultset quando o modo inclui o Dashboard.';
PRINT 'O relatório HTML é o último resultset quando @GerarHTML = 1.';
PRINT '=================================================================';

DROP TABLE IF EXISTS #Dashboard, #AssessmentSummary, #LogSpace, #Secoes, #ColetaErros, #Achados,
    #D01_Instancia, #D02_Hardware, #D02_HostInfo, #D02_Servicos, #D03_ClusterNodes, #D03_ClusterProps,
    #D04_AGs, #D04_Replicas, #D04_Databases, #D04_Cluster, #D05_LinkedServers, #D06_Endpoints,
    #D07_Databases, #D08_Arquivos, #D08_Volumes, #D08_Drives, #D09_TempdbArquivos, #D09_VersionStore,
    #D10_Backups, #D11_Jobs, #D11_Alertas, #D12_Logins, #D13_Config, #D13_ConfigPendente,
    #D14_ProcessMemory, #D14_SysMemory, #D14_PLE, #D14_CPU, #D15_IOLatencia, #D15_IOPorDB,
    #IOWarningResults, #D16_SuspectPages, #D16_TraceFlags, #D17_BufferPool, #D17_VLF, #D18_Waits,
    #D19_Permissoes, #D19_ServerRoles, #D19_PoliticaSenha, #D20_Objetos, #V10_FragBase, #D20_Fragmentacao,
    #D20_FragResumo, #D20_MissingIndexes, #D21_TopDuracao, #D21_TopCPU, #D22_Sessoes, #D23_BackupHist,
    #D23_BackupLog, #D24_JobAgenda, #V11_LSJob, #D25_LSPrimario, #D25_LSSecundario, #D25_LSMonitor,
    #D25_LSJobs, #D25_LSErros, #V11_HtmlDet;
GO
