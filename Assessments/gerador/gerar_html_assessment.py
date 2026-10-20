#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera o relatório HTML do SQL Server Assessment (Solo Network) a partir do
arquivo .rpt produzido pelo SSMS (Resultados em arquivo) ao executar o
SQL_Server_Assessment_Profissional_v11.sql no cliente.

O HTML segue o layout do relatório v11 (guias, resumo executivo, checklist)
com o papel timbrado Solo embutido (template/cabecalho.html e
template/rodape_solo.b64). Não depende de bibliotecas externas.

Uso:
    python gerar_html_assessment.py ARQUIVO.rpt --cliente "Nome do Cliente"
    python gerar_html_assessment.py ARQUIVO.rpt --cliente "X" --saida relatorio.html

Veja Assessments/PROCESSO.md para o mapeamento resultset -> seção do HTML.
"""
import argparse
import html as _html
import os
import re
import sys
from collections import OrderedDict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

AQUI = os.path.dirname(os.path.abspath(__file__))
MAX_LINHAS = 500  # igual ao @HtmlMaxLinhas do script

# ---------------------------------------------------------------------------
# Catálogo de seções (espelha o #Secoes do script v11)
#   chave: primeiras colunas do resultset (minúsculas)
#   valor: (grupo, título, descrição, colunas largas)
# ---------------------------------------------------------------------------
SECOES = [
    (('machinename', 'servername'), '01', 'Propriedades da instância',
     'SERVERPROPERTY: versão, edição, collation, caminhos padrão e recursos instalados.', ''),
    (('logicalcpucount',), '02', 'Hardware (CPU, memória, NUMA)', 'sys.dm_os_sys_info.', ''),
    (('host_platform',), '02', 'Sistema operacional', 'sys.dm_os_host_info.', ''),
    (('servicename',), '02', 'Serviços do SQL Server',
     'Contas de serviço, tipo de inicialização e Instant File Initialization.', 'filename'),
    (('nodename',), '03', 'Nós do cluster (FCI)', 'sys.dm_os_cluster_nodes.', ''),
    (('verboselogging',), '03', 'Propriedades do cluster (FCI)', 'sys.dm_os_cluster_properties.', ''),
    (('agname', 'group_id'), '04', 'Availability Groups', 'Configuração de cada AG.', ''),
    (('agname', 'replica_server_name', 'endpoint_url'), '04', 'Réplicas dos AGs',
     'Modo de disponibilidade, failover, seeding e saúde de cada réplica.', ''),
    (('agname', 'replica_server_name', 'database_name'), '04', 'Databases nos AGs',
     'Estado de sincronização, filas de envio/redo (KB) e últimos tempos por database/réplica.', ''),
    (('cluster_name', 'quorum_type_desc'), '04', 'Cluster WSFC / quórum', 'sys.dm_hadr_cluster.', ''),
    (('name', 'product'), '05', 'Linked Servers', 'Provider, data source e opções de RPC/data access.', ''),
    (('name', 'endpoint_id'), '06', 'Endpoints', 'Endpoints TCP, HADR e Service Broker.', ''),
    (('database_id', 'databasename'), '07', 'Databases',
     'Estado, recovery model, compatibilidade, opções e tamanho de cada database.', ''),
    (('databasename', 'file_id', 'logicalname'), '08', 'Arquivos de dados e log',
     'Localização, tamanho e configuração de crescimento de cada arquivo.', 'physical_name'),
    (('volume_mount_point',), '08', 'Volumes com arquivos do SQL Server', 'sys.dm_os_volume_stats.', ''),
    (('fixed_drive_path',), '08', 'Unidades fixas do servidor', 'sys.dm_os_enumerate_fixed_drives (Windows).', ''),
    (('name', 'physical_name'), '09', 'TempDB - arquivos',
     'Quantidade, tamanho e crescimento dos arquivos do TempDB.', 'physical_name'),
    (('databasename', 'reserved_page_count'), '09', 'TempDB - version store por database',
     'Espaço de version store (RCSI/SI, AG legível, triggers, online index).', ''),
    (('databasename', 'recovery_model_desc'), '10', 'Backups (últimos 30 dias) e último CHECKDB',
     'Datas vazias = sem backup nos últimos 30 dias. LastGoodCheckDB 01/01/1900 = CHECKDB nunca executado com sucesso.', ''),
    (('jobname',), '11', 'Jobs do SQL Agent',
     'Owner, categoria e resultado da última execução. LastRunDuration no formato HHMMSS.', 'LastRunMessage'),
    (('name', 'event_source'), '11', 'Alertas do SQL Agent', 'msdb.dbo.sysalerts.', ''),
    (('loginname',), '12', 'Logins',
     'Logins SQL, usuários e grupos Windows. IsSysadmin = 1 indica membro da role sysadmin.', ''),
    (('name', 'configuredvalue'), '13', 'Configurações relevantes (sp_configure)',
     'Principais configurações de desempenho e superfície de ataque.', 'description'),
    (('configuration_id', 'name'), '13', 'Configurações pendentes (RECONFIGURE/restart)',
     'Configurações cujo valor configurado difere do valor em uso.', 'description'),
    (('sqlmemoryusedmb',), '14', 'Memória do processo SQL Server', 'sys.dm_os_process_memory.', ''),
    (('physicalmemorymb', 'availablephysicalmemorymb'), '14', 'Memória do sistema operacional',
     'sys.dm_os_sys_memory.', ''),
    (('objectname', 'instance_name'), '14', 'Page Life Expectancy por nó NUMA',
     'Valor pontual; avaliar em conjunto com a tendência.', ''),
    (('sqlserverprocesscpu',), '14', 'Histórico de CPU (ring buffer, 1 amostra/min)',
     'Uso de CPU pelo SQL Server e por outros processos nos últimos minutos.', ''),
    (('databasename', 'file_id', 'physical_name'), '15', 'Latência de I/O por arquivo',
     'Médias acumuladas desde o último restart. Referência: dados < 20 ms, log < 5 ms.', 'physical_name'),
    (('iorank',), '15', 'Volume de I/O por database', 'Distribuição do I/O acumulado desde o último restart.', ''),
    (('logdate', 'processinfo'), '15', 'Avisos de I/O > 15 s no ErrorLog',
     'Ocorrências de "taking longer than 15 seconds" nos seis ErrorLogs mais recentes.', 'LogText'),
    (('databasename', 'file_id', 'page_id'), '16', 'Suspect pages',
     'event_type 1-3 = ativos; 4 = restaurado; 5 = reparado; 7 = desalocado pelo DBCC.', 'physical_name'),
    (('traceflag',), '16', 'Trace flags globais ativos', 'DBCC TRACESTATUS(-1).', ''),
    (('bufferpoolrank',), '17', 'Buffer pool por database', 'Distribuição das páginas em cache no momento da coleta.', ''),
    (('databasename', 'logsizemb'), '17', 'Uso do log de transações', 'DBCC SQLPERF(LOGSPACE).', ''),
    (('databasename', 'file_id', 'vlfcount'), '17', 'VLFs por arquivo de log',
     'Referência: acima de 1.000 VLFs merece revisão do crescimento do log.', ''),
    (('wait_type',), '18', 'Top 20 wait stats',
     'Esperas acumuladas desde o último restart (ou limpeza), excluindo esperas benignas/ociosas. PctTotal é relativo às esperas filtradas.', ''),
    (('pollogin',), '19', 'Política de senha dos logins SQL',
     'CHECK_POLICY/CHECK_EXPIRATION, última troca de senha, bloqueio e verificação de senha em branco ou igual ao login (PWDCOMPARE).', ''),
    (('srvrolename',), '19', 'Membros de server roles',
     'Todas as server roles (fixas e criadas pelo usuário). securityadmin deve ser tratada como sysadmin.', ''),
    (('permprincipal',), '19', 'Permissões explícitas de servidor',
     'sys.server_permissions, omitindo o GRANT CONNECT SQL padrão e as permissões padrão do role public.', ''),
    (('objdatabase',), '20', 'Objetos por database',
     'Objetos de usuário (is_ms_shipped = 0): tabelas, views, procedures, functions, triggers, índices, sinônimos e schemas.', ''),
    (('fragresdatabase',), '20', 'Resumo de fragmentação por database',
     'Índices com {minpag} páginas ou mais, modo LIMITED. Críticos >= 70%, moderados 30-70%.', ''),
    (('fragdatabase',), '20', 'Índices fragmentados (>= 30%)',
     'Até 500 índices, do mais fragmentado para o menos. FragIndexId 0 = heap (ALTER TABLE ... REBUILD).', ''),
    (('midatabase',), '20', 'Missing indexes (top 25)',
     'Sugestões do otimizador desde o último restart. Não consideram índices existentes nem o custo de escrita: consolidar antes de criar.',
     'MIEquality,MIInequality,MIIncluded,MICreateStatement'),
    (('tqdatabase',), '21', 'Top 20 queries por duração média',
     'sys.dm_exec_query_stats: somente planos em cache, com ao menos 2 execuções. Tempos em milissegundos.', 'TQSqlText'),
    (('tcdatabase',), '21', 'Top 20 queries por CPU total',
     'Consumo acumulado de CPU desde a compilação do plano. Tempos em milissegundos.', 'TCSqlText'),
    (('sesssessionid',), '22', 'Sessões de usuário no momento da coleta',
     'Fotografia do momento da execução do assessment (exclui a própria sessão). SessBlockedBy <> 0 = sessão bloqueada.', 'SessSqlText'),
    (('bkdatabase',), '23', 'Histórico de backups Full/Diferencial (7 dias)',
     'msdb desta instância. BkDeviceType 2 = disco, 7 = dispositivo virtual (ferramenta externa), 9 = URL. BkCompressionRatio ~1 = sem compressão.', 'BkDevice'),
    (('logbkdatabase',), '23', 'Backups de log - resumo diário (7 dias)',
     'Quantidade, volume e janela dos backups de log por database e dia.', ''),
    (('schjobname',), '24', 'Agendamentos dos jobs',
     'Frequência, dias, horário e próxima execução de cada agendamento (exceto assinaturas do Reporting Services).', ''),
    (('lspdatabase',), '25', 'Log shipping - databases primárias',
     'Configuração do envio de logs: pasta local e compartilhamento de backup, retenção (min), compressão, servidor monitor, secundários, último backup de log e minutos decorridos x limite de alerta.', ''),
    (('lssdatabase',), '25', 'Log shipping - databases secundárias',
     'Origem (primário), pastas de cópia, modo de restore (NORECOVERY ou STANDBY), atraso configurado, última cópia, último restore, latência e limite de alerta. Limite efetivo = LSSLimiteMin + LSSAtrasoRestoreMin.', ''),
    (('lsmpapel',), '25', 'Log shipping - databases monitoradas (servidor monitor)',
     'Databases de outras instâncias que usam esta instância como servidor monitor. Tempos calculados em UTC.', ''),
    (('lsjfuncao',), '25', 'Log shipping - jobs',
     'Jobs LSBackup (primário), LSCopy e LSRestore (secundário) e LSAlert (monitor): habilitado, última execução, resultado e próxima execução.', ''),
    (('lsedata',), '25', 'Log shipping - erros recentes (7 dias)',
     'Últimas 50 mensagens de msdb.dbo.log_shipping_monitor_error_detail.', 'LSEMensagem'),
]

# Seções apenas informativas (resultset de uma coluna "Information")
INFORMATIVAS = [
    (r'FCI', '03', 'Cluster / FCI'),
    (r'Always On', '04', 'Always On'),
    (r'fragmenta', '20', 'Índices fragmentados'),
    (r'[Ll]og shipping', '25', 'Log shipping'),
]

# Colunas bit (no .rpt aparecem como 0/1; no HTML viram Sim/Não)
COLUNAS_BIT = {c.lower() for c in '''
is_auto_close_on is_auto_shrink_on is_auto_create_stats_on is_auto_update_stats_on is_auto_update_stats_async_on
is_parameterization_forced is_read_committed_snapshot_on is_cdc_enabled is_encrypted is_query_store_on
is_percent_growth is_dynamic is_advanced is_disabled is_dynamic_port is_compressed supports_compression
supports_sparse_files supports_alternate_streams process_physical_memory_low process_virtual_memory_low
is_linked is_data_access_enabled is_rpc_out_enabled is_remote_login_enabled is_system is_encryption_enabled
is_local is_primary_replica db_failover dtc_support basic_features is_current_owner JobEnabled
SrvRoleMemberDisabled PolDisabled PolPolicyChecked PolExpirationChecked PolIsLocked PolIsExpired
PolBlankPassword PolPasswordEqualsLogin BkCopyOnly BkSnapshot SchEnabled LSPAlertaHabilitado
LSSAlertaHabilitado LSSDesconectaUsuarios LSMAlertaHabilitado LSJHabilitado
'''.split()}

# Colunas texto que parecem número/data no .rpt
COLUNAS_TEXTO = {c.lower() for c in 'instance_name LogBkDay host_release'.split()}
SECOES_CRUAS = {'Propriedades da instância'}  # sql_variant: exibido sem formatação

LST_OK = {'OK', 'HEALTHY', 'SYNCHRONIZED', 'ONLINE', 'CONNECTED', 'STARTED', 'Running', 'Sucesso'}
LST_WARN = {'ATENÇÃO', 'PARTIALLY_HEALTHY', 'RESTORING', 'RECOVERING', 'Stopped', 'Paused', 'Cancelado', 'Nova tentativa'}
LST_CRIT = {'CRÍTICO', 'NOT_HEALTHY', 'SUSPECT', 'EMERGENCY', 'OFFLINE', 'RECOVERY_PENDING', 'DISCONNECTED',
            'NOT SYNCHRONIZING', 'NOT_SYNCHRONIZING', 'Falha'}

GUIAS = [
    (1, 'geral', 'Visão Geral', 'Visão Geral', None),
    (2, 'atencao', 'Pontos de Atenção', 'Atenção', None),
    (3, 'databases', 'Databases', 'Databases',
     'Configuração, tamanho e estado das databases, objetos, TempDB e arquivos de log (VLFs).'),
    (4, 'performance', 'Performance', 'Performance',
     'Uso de CPU, memória, I/O, buffer pool, esperas, índices, queries mais pesadas e configurações que afetam o desempenho.'),
    (5, 'backups', 'Backups e Integridade', 'Backups', 'Histórico de backups, último CHECKDB e páginas suspeitas.'),
    (6, 'seguranca', 'Segurança', 'Segurança',
     'Logins, membros de server roles, permissões, política de senha e configurações de segurança.'),
    (7, 'jobs', 'Jobs e Alertas', 'Jobs',
     'Jobs do SQL Server Agent, agendamentos, resultado da última execução e alertas configurados.'),
    (8, 'armazenamento', 'Armazenamento', 'Armazenamento',
     'Volumes, maiores databases e localização dos arquivos de dados e log.'),
    (9, 'conexoes', 'Conexões', 'Conexões', 'Linked servers, endpoints e sessões de usuário no momento da coleta.'),
    (10, 'ha', 'Alta Disponibilidade', 'Alta Disponibilidade', 'Cluster (FCI), Always On Availability Groups e Log Shipping.'),
    (11, 'checklist', 'Checklist', 'Checklist', None),
]

GUIA_ACHADO = {'01': 'geral', '02': 'geral', '03': 'ha', '04': 'ha', '23': 'ha', '05': 'conexoes', '06': 'conexoes',
               '21': 'conexoes', '07': 'databases', '09': 'databases', '17': 'databases', '08': 'armazenamento',
               '10': 'backups', '16': 'backups', '22': 'backups', '11': 'jobs', '12': 'seguranca', '18': 'seguranca',
               '13': 'performance', '14': 'performance', '15': 'performance', '19': 'performance', '20': 'performance'}

GUIA_SECAO = {'01': 'geral', '02': 'geral', '03': 'ha', '04': 'ha', '25': 'ha', '05': 'conexoes', '06': 'conexoes',
              '22': 'conexoes', '07': 'databases', '09': 'databases', '17': 'databases', '08': 'armazenamento',
              '10': 'backups', '16': 'backups', '23': 'backups', '11': 'jobs', '24': 'jobs', '12': 'seguranca',
              '19': 'seguranca', '13': 'performance', '14': 'performance', '15': 'performance', '18': 'performance',
              '20': 'performance', '21': 'performance'}

TOOLS = ('<div class="tools"><button onclick="tg(true)" type="button">Expandir tudo</button>'
         '<button onclick="tg(false)" type="button">Recolher tudo</button></div>')
CABEC_ACHADOS = '<div class="tw"><table><thead><tr><th>Indicador</th><th>Valor</th><th>Status</th><th>Observação</th></tr></thead><tbody>'


def esc(s):
    return _html.escape('' if s is None else str(s), quote=False)


# ---------------------------------------------------------------------------
# Leitura do .rpt
# ---------------------------------------------------------------------------
class Resultset:
    def __init__(self, colunas, linhas, linha_arquivo):
        self.colunas = colunas
        self.linhas = linhas  # lista de listas (None = NULL)
        self.linha_arquivo = linha_arquivo

    def col(self, nome):
        n = nome.lower()
        for i, c in enumerate(self.colunas):
            if c.lower() == n:
                return i
        return -1

    def valores(self, nome):
        i = self.col(nome)
        return [] if i < 0 else [r[i] for r in self.linhas]

    def primeiro(self, nome):
        v = self.valores(nome)
        return v[0] if v else None


def ler_texto(caminho):
    dados = open(caminho, 'rb').read()
    if dados[:2] in (b'\xff\xfe', b'\xfe\xff'):
        return dados.decode('utf-16')
    if dados[:3] == b'\xef\xbb\xbf':
        return dados[3:].decode('utf-8')
    # UTF-16 sem BOM (muitos bytes nulos)
    if dados[1:200:2].count(0) > 50:
        return dados.decode('utf-16-le')
    try:
        return dados.decode('utf-8')
    except UnicodeDecodeError:
        return dados.decode('cp1252')


RE_TRACOS = re.compile(r'^-+( +-+)*\s*$')
RE_TRACOS_TAB = re.compile(r'^-+(\t-+)*\s*$')
COLUNAS_UNICAS = {'Information', 'Relatorio_HTML'}  # resultsets de uma coluna conhecidos
RE_AFETADAS = re.compile(r'^\(\d+ (rows?|linhas?|linha\(s\)|row\(s\)) (affected|afetadas?)\)\s*$', re.I)


def ler_rpt(texto):
    linhas = texto.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    resultsets, mensagens = [], []
    i, n = 0, len(linhas)
    while i < n:
        lin = linhas[i]
        prox = linhas[i + 1] if i + 1 < n else ''
        if lin.strip() and (RE_TRACOS.match(prox) or RE_TRACOS_TAB.match(prox)) and prox.strip():
            if '\t' in prox:
                cols = [c.strip() for c in lin.split('\t')]
                spans = None
            else:
                spans = [(m.start(), m.end()) for m in re.finditer(r'-+', prox)]
                cols = [lin[a:b].strip() if k < len(spans) - 1 else lin[a:].strip()
                        for k, (a, b) in enumerate(spans)]
            j = i + 2
            dados = []
            while j < n:
                l = linhas[j]
                if not l.strip() or RE_AFETADAS.match(l):
                    break
                if j + 1 < n and RE_TRACOS.match(linhas[j + 1]) and linhas[j + 1].strip() \
                        and len(linhas[j + 1]) >= len(l.rstrip()) - 2:
                    break
                if spans is None:
                    vals = l.split('\t')
                    j += 1
                else:
                    ini_ult, largura = spans[-1][0], spans[-1][1]

                    def na_grade(x):
                        return len(x) >= ini_ult and all(x[a - 1] == ' ' for a, _ in spans[1:])

                    # valor com quebra de linha (ex.: mensagem de job): o SSMS imprime a quebra
                    # no meio da linha; junta as linhas seguintes que não se encaixam na grade
                    while len(l) < largura and j + 1 < n and linhas[j + 1].strip() \
                            and not na_grade(linhas[j + 1]) and not RE_AFETADAS.match(linhas[j + 1]):
                        j += 1
                        l = l + '\n' + linhas[j]
                    vals = []
                    for k, (a, b) in enumerate(spans):
                        fim = b if k < len(spans) - 1 else len(l)
                        vals.append(l[a:fim])
                    j += 1
                vals = [None if v.strip() == 'NULL' else v.strip() for v in vals]
                vals += [None] * (len(cols) - len(vals))
                dados.append(vals[:len(cols)])
            resultsets.append(Resultset(cols, dados, i + 1))
            i = j
            continue
        # "Resultados em arquivo" com delimitador tab: cabeçalho sem linha de traços,
        # linhas até a próxima linha em branco
        tabs = lin.count('\t')
        if (tabs and (not prox.strip() or prox.count('\t') >= tabs)) or (lin.strip() in COLUNAS_UNICAS and prox.strip()):
            cols = [c.strip() for c in lin.split('\t')]
            j = i + 1
            dados = []
            while j < n and linhas[j].strip() and not RE_AFETADAS.match(linhas[j]):
                l = linhas[j]
                # valor com quebra de linha: junta até completar as colunas
                while l.count('\t') < tabs and j + 1 < n and linhas[j + 1].strip():
                    j += 1
                    l = l + '\n' + linhas[j]
                vals = [None if v.strip() == 'NULL' else v.strip() for v in l.split('\t')]
                vals += [None] * (len(cols) - len(vals))
                dados.append(vals[:len(cols)])
                j += 1
            resultsets.append(Resultset(cols, dados, i + 1))
            i = j
            continue
        if lin.strip():
            mensagens.append(lin.rstrip())
        i += 1
    return resultsets, mensagens


# ---------------------------------------------------------------------------
# Formatação (igual ao CONVERT/FORMAT do script)
# ---------------------------------------------------------------------------
RE_INT = re.compile(r'^-?\d+$')
RE_DEC = re.compile(r'^-?(\d+\.\d*|\.\d+|\d+(\.\d+)?[eE][-+]?\d+)$')
RE_DT = re.compile(r'^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?(\.\d+)?( ?[+-]\d{2}:\d{2})?$')


def fmt_num(valor, casas):
    try:
        d = Decimal(str(valor))
    except InvalidOperation:
        return str(valor)
    q = Decimal(1).scaleb(-casas)
    d = d.quantize(q, rounding=ROUND_HALF_UP)
    s = f'{abs(d):,.{casas}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return ('-' if d < 0 else '') + s


def fmt_int_milhar(n):
    return f'{int(n):,}'.replace(',', '.')


def parse_dt(v):
    m = RE_DT.match(v or '')
    if not m:
        return None
    y, mo, d, h, mi, s = (int(x) if x else 0 for x in m.groups()[:6])
    return datetime(y, mo, d, h, mi, s)


def fmt_dt(dt, segundos=True):
    return dt.strftime('%d/%m/%Y %H:%M:%S' if segundos else '%d/%m/%Y %H:%M')


def tipos_colunas(rs):
    tipos = []
    for k, c in enumerate(rs.colunas):
        vals = [r[k] for r in rs.linhas if r[k] not in (None, '')]
        cl = c.lower()
        if cl in COLUNAS_TEXTO or not vals:
            t = 'txt'
        elif cl in COLUNAS_BIT and all(v in ('0', '1') for v in vals):
            t = 'bit'
        elif all(RE_INT.match(v) for v in vals):
            t = 'int'
        elif all(RE_INT.match(v) or RE_DEC.match(v) for v in vals):
            t = 'dec'
        elif all(RE_DT.match(v) for v in vals):
            t = 'dt'
        else:
            t = 'txt'
        tipos.append(t)
    return tipos


def fmt_valor(v, t):
    if v is None:
        return ''
    if t == 'bit':
        return 'Sim' if v == '1' else 'Não'
    if t == 'dec':
        return fmt_num(v, 2)
    if t == 'dt':
        dt = parse_dt(v)
        return fmt_dt(dt) if dt else v
    return v


def celula(v, t, larga):
    s = esc(v)
    if t in ('int', 'dec'):
        return '<td class="n">' + (s or '—') + '</td>'
    if v in LST_OK:
        return '<td><span class="b ok">' + s + '</span></td>'
    if v in LST_WARN:
        return '<td><span class="b warn">' + s + '</span></td>'
    if v in LST_CRIT:
        return '<td><span class="b crit">' + s + '</span></td>'
    return ('<td class="w">' if larga else '<td>') + (s or '—') + '</td>'


# ---------------------------------------------------------------------------
# Montagem
# ---------------------------------------------------------------------------
def identificar(rs):
    cols = tuple(c.lower() for c in rs.colunas)
    melhor, tam = None, 0
    for chave, *meta in SECOES:
        if cols[:len(chave)] == chave and len(chave) > tam:
            melhor, tam = meta, len(chave)
    return melhor


def contagem(n):
    if n == 0:
        return 'nenhum registro'
    if n == 1:
        return '1 registro'
    return fmt_int_milhar(n) + ' registros'


def guia_secao(grupo, titulo):
    t = titulo.lower()
    if 'buffer pool' in t or 'trace flag' in t:
        return 'performance'
    if 'log de transações' in t or 'vlf' in t:
        return 'databases'
    if grupo == '20' and titulo.startswith('Objetos'):
        return 'databases'
    return GUIA_SECAO.get(grupo, 'geral')


class Relatorio:
    def __init__(self, rpt, cliente, nome_arquivo):
        self.cliente = cliente
        self.nome_arquivo = nome_arquivo
        texto = ler_texto(rpt)
        self.texto = texto
        self.resultsets, self.mensagens = ler_rpt(texto)
        self.dashboard = None
        self.summary = None
        self.secoes = []  # dicts: grupo, titulo, desc, largas, rs (ou nota)
        self.por_titulo = {}
        self.nao_mapeados = []
        self.html_pronto = ''
        self._classificar()

    # -- classificação dos resultsets ------------------------------------
    def _classificar(self):
        for rs in self.resultsets:
            cols = [c.lower() for c in rs.colunas]
            if cols[:6] == ['id', 'categoria', 'indicador', 'valor', 'status', 'observacao']:
                cats = [r[1] or '' for r in rs.linhas]
                if cats and all(re.match(r'^\d\d - ', c) for c in cats):
                    self.summary = rs
                else:
                    self.dashboard = rs
                continue
            if cols and cols[0] in ('relatorio_html', 'parte'):
                k = rs.col('Relatorio_HTML') if rs.col('Relatorio_HTML') >= 0 else rs.col('Relatorio_HTML_Texto')
                self.html_pronto = ''.join(r[k] or '' for r in rs.linhas)
                continue
            if cols == ['information']:
                nota = rs.linhas[0][0] if rs.linhas else ''
                grupo, titulo = '??', 'Informação'
                for padrao, g, t in INFORMATIVAS:
                    if re.search(padrao, nota or ''):
                        grupo, titulo = g, t
                        break
                self.secoes.append(dict(grupo=grupo, titulo=titulo, nota=nota, rs=None))
                continue
            meta = identificar(rs)
            if meta is None:
                self.nao_mapeados.append(rs)
                meta = ('??', 'Resultset não mapeado (' + ', '.join(rs.colunas[:3]) + ')', None, '')
            grupo, titulo, desc, largas = meta
            if desc and '{minpag}' in desc:
                mp = rs.primeiro('FragResMinPaginas') or '1000'
                desc = desc.replace('{minpag}', mp)
            sec = dict(grupo=grupo, titulo=titulo, desc=desc, largas=largas, rs=rs)
            self.secoes.append(sec)
            self.por_titulo.setdefault(titulo, rs)

    def rs(self, titulo):
        return self.por_titulo.get(titulo)

    # -- dados auxiliares ------------------------------------------------
    def dash(self, indicador):
        for fonte in (self.dashboard, self.summary):
            if fonte is None:
                continue
            for r in fonte.linhas:
                if (r[2] or '').strip().lower() == indicador.lower():
                    return r[3]
        return None

    def prop(self, nome):
        rs = self.rs('Propriedades da instância')
        return rs.primeiro(nome) if rs else None

    def hw(self, nome):
        rs = self.rs('Hardware (CPU, memória, NUMA)')
        return rs.primeiro(nome) if rs else None

    def versao_script(self):
        m = re.search(r'ASSESSMENT / DISCOVERY - V(\d+)', self.texto, re.I)
        if m:
            return 'v' + m.group(1)
        grupos = {s['grupo'] for s in self.secoes}
        if '25' in grupos:
            return 'v11'
        if grupos & {'19', '20', '21', '22', '23', '24'}:
            return 'v10'
        if '18' in grupos:
            return 'v9'
        return 'v8'

    def data_coleta(self):
        m = re.search(r'Data: (\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', self.texto)
        if m:
            return parse_dt(m.group(1))
        cands = []
        for titulo, col in (('Page Life Expectancy por nó NUMA', 'CollectionTime'),
                            ('Histórico de CPU (ring buffer, 1 amostra/min)', 'EventTime')):
            rs = self.rs(titulo)
            if rs:
                cands += [parse_dt(v) for v in rs.valores(col) if parse_dt(v)]
            if cands:
                return max(cands)
        return datetime.fromtimestamp(os.path.getmtime(self.nome_arquivo))

    def achados(self):
        lst = []
        if self.summary is not None:
            for r in self.summary.linhas:
                lst.append(dict(ordem=int(r[0]) if RE_INT.match(r[0] or '') else len(lst) + 1,
                                cat=r[1] or '', ind=r[2] or '', val=r[3], st=(r[4] or 'INFO'), obs=r[5]))
        elif self.dashboard is not None:
            for r in self.dashboard.linhas:
                if r[0] in ('26', '27'):
                    continue
                lst.append(dict(ordem=int(r[0]), cat=r[1] or '', ind=r[2] or '', val=r[3], st=r[4] or 'INFO', obs=r[5]))
        for a in lst:
            a['guia'] = 'seguranca' if a['ind'] == 'xp_cmdshell' else GUIA_ACHADO.get(a['cat'][:2], 'geral')
        return lst

    # -- blocos HTML -----------------------------------------------------
    def html_secao(self, idx, sec):
        cab = (f'<details id="s{idx}" open="open"><summary><span class="g">{esc(sec["grupo"])}</span>'
               f'{esc(sec["titulo"])}')
        if sec['rs'] is None:
            return cab + '<span class="c">informativo</span></summary><div class="note">' + esc(sec['nota'] or '-') + '</div></details>'
        rs = sec['rs']
        linhas = list(rs.linhas)
        total = len(linhas)
        nota_extra = ''
        if sec['titulo'] == 'Jobs do SQL Agent':
            ic = rs.col('CategoryName')
            if ic >= 0 and any(r[ic] == 'Report Server' for r in linhas):
                linhas = [r for r in linhas if r[ic] != 'Report Server'] + [r for r in linhas if r[ic] == 'Report Server']
                nota_extra = ' (jobs do Report Server listados por último)'
        tipos = tipos_colunas(rs)
        if sec['titulo'] in SECOES_CRUAS:
            tipos = ['txt'] * len(tipos)
        largas = {c.strip().lower() for c in (sec['largas'] or '').split(',') if c.strip()}
        h = cab + f'<span class="c">{contagem(total)}</span></summary>'
        if sec['desc']:
            h += '<p class="desc">' + esc(sec['desc']) + '</p>'
        if total == 1 and len(rs.colunas) > 6:
            h += '<div class="tw"><table class="data kv"><thead><tr><th>Propriedade</th><th>Valor</th></tr></thead><tbody>'
            for c, v, t in zip(rs.colunas, linhas[0], tipos):
                h += '<tr><th>' + esc(c) + '</th><td>' + (esc(fmt_valor(v, t)) or '—') + '</td></tr>'
            h += '</tbody></table></div>'
        else:
            h += '<div class="tw"><table class="data"><thead><tr>' + ''.join('<th>' + esc(c) + '</th>' for c in rs.colunas) + '</tr></thead><tbody>'
            if total == 0:
                h += f'<tr><td class="empty" colspan="{len(rs.colunas)}">Nenhum registro encontrado.</td></tr>'
            for r in linhas[:MAX_LINHAS]:
                h += '<tr>' + ''.join(celula(fmt_valor(v, t), t, c.lower() in largas)
                                      for c, v, t in zip(rs.colunas, r, tipos)) + '</tr>'
            h += '</tbody></table></div>'
        if total > MAX_LINHAS:
            h += (f'<div class="note">Exibindo {MAX_LINHAS} de {fmt_int_milhar(total)} registros{nota_extra}. '
                  'A lista completa está no arquivo .rpt original.</div>')
        return h + '</details>'

    def linha_achado(self, a):
        cls = {'OK': 'ok', 'ATENÇÃO': 'warn', 'CRÍTICO': 'crit'}.get(a['st'], 'info')
        return ('<tr><td>' + esc(a['ind']) + '</td><td class="n">' + (esc(a['val']) or '—') + '</td>'
                f'<td><span class="b {cls}">' + esc(a['st']) + '</span></td><td class="obs">' + (esc(a['obs']) or '—') + '</td></tr>')

    def capacidade(self):
        vol = self.rs('Volumes com arquivos do SQL Server')
        partes = []
        if vol:
            for r in sorted(vol.linhas, key=lambda r: r[vol.col('volume_mount_point')] or ''):
                try:
                    tot = Decimal(r[vol.col('TotalGB')])
                    av = Decimal(r[vol.col('AvailableGB')])
                except (InvalidOperation, TypeError):
                    continue
                if tot <= 0:
                    continue
                usedpct = int((100 - av * 100 / tot).quantize(Decimal(1), ROUND_HALF_UP))
                freepct = int((av * 100 / tot).quantize(Decimal(1), ROUND_HALF_UP))
                cls = 'crit' if freepct < 10 else 'warn' if freepct < 15 else 'ok'
                partes.append(f'<span class="n">{esc(r[vol.col("volume_mount_point")])}</span><div class="bar"><i class="{cls}" '
                              f'style="width:{usedpct}%">{usedpct}</i></div><span class="t">{fmt_num(tot - av, 1)} / '
                              f'{fmt_num(tot, 1)} GB · {freepct}% livre</span>')
        dbs = self.rs('Databases')
        barras = []
        if dbs:
            tam = []
            for r in dbs.linhas:
                try:
                    mb = Decimal(r[dbs.col('DataSizeMB')] or 0) + Decimal(r[dbs.col('LogSizeMB')] or 0)
                except InvalidOperation:
                    continue
                tam.append((r[dbs.col('DatabaseName')], mb / 1024))
            tam = sorted(tam, key=lambda x: -x[1])[:10]
            mx = tam[0][1] if tam else 0
            for nome, gb in tam:
                p = gb * 100 / mx if mx else 0
                pct = 1 if p < 1 else int(p.quantize(Decimal(1), ROUND_HALF_UP))
                barras.append(f'<span class="n">{esc(nome)}</span><div class="bar"><i class="db" style="width:{pct}%">{pct}</i></div>'
                              f'<span class="t">{fmt_num(gb, 2 if gb < 10 else 1)} GB</span>')
        return ('<div class="grid2"><div class="card"><h3>Volumes com arquivos do SQL Server</h3>'
                + ('<div class="hb">' + ''.join(partes) + '</div>' if partes else '<p class="empty">Informação de volumes indisponível.</p>')
                + '</div><div class="card"><h3>Maiores databases (dados + log)</h3>'
                + ('<div class="hb">' + ''.join(barras) + '</div>' if barras else '<p class="empty">Informação indisponível.</p>')
                + '</div></div>')

    def cpu(self):
        rs = self.rs('Histórico de CPU (ring buffer, 1 amostra/min)')
        pts = []
        if rs:
            for r in rs.linhas:
                dt = parse_dt(r[rs.col('EventTime')])
                try:
                    sql = int(r[rs.col('SQLServerProcessCPU')] or 0)
                    oth = r[rs.col('OtherProcessCPU')]
                    oth = int(oth) if oth not in (None, '') else max(0, 100 - int(r[rs.col('SystemIdleCPU')] or 0) - sql)
                except ValueError:
                    continue
                if dt:
                    pts.append((dt, sql, max(0, oth)))
        pts.sort(key=lambda p: p[0])
        n = len(pts)
        if n <= 1:
            return '<div class="card cpu"><h3>Uso de CPU</h3><p class="empty">Histórico de CPU indisponível (ring buffer vazio ou sem permissão).</p></div>'

        def y(v):
            return fmt_ponto(10 + Decimal(140) * (1 - Decimal(v) / 100))

        xs = [fmt_ponto(Decimal(40) + Decimal(i) * 960 / (n - 1)) for i in range(n)]
        p_sql = ' '.join(f'{x},{y(p[1])}' for x, p in zip(xs, pts))
        p_oth = ' '.join(f'{x},{y(p[2])}' for x, p in zip(xs, pts))
        media = int(sum(p[1] for p in pts) / n)
        pico = max(p[1] for p in pts)
        grade = ''.join(f'<line class="grid" x1="40" x2="1000" y1="{yy}" y2="{yy}"></line>'
                        f'<text class="ax" text-anchor="end" x="32" y="{yy + 4}">{lbl}</text>'
                        for yy, lbl in ((10, '100%'), (45, '75%'), (80, '50%'), (115, '25%'), (150, '0%')))
        return (f'<div class="card cpu"><h3>Uso de CPU · últimos {n} minutos</h3>'
                f'<p class="sub">SQL Server: média {media}% · pico {pico}% (amostras por minuto do ring buffer do scheduler monitor)</p>'
                f'<svg viewbox="0 0 1000 180">{grade}<polyline class="l-oth" points="{p_oth}"></polyline>'
                f'<polyline class="l-sql" points="{p_sql}"></polyline>'
                f'<text class="ax" x="40" y="174">{fmt_dt(pts[0][0], False)}</text>'
                f'<text class="ax" text-anchor="end" x="1000" y="174">{fmt_dt(pts[-1][0], False)}</text>'
                '</svg><div class="legend"><span>SQL Server</span><span class="o">Outros processos</span></div></div>')

    # -- documento -------------------------------------------------------
    def gerar(self):
        achados = self.achados()
        ver = self.versao_script()
        coleta = self.data_coleta()
        data_coleta = fmt_dt(coleta, False)
        servidor = self.prop('ServerName') or self.dash('Servidor') or 'N/D'

        n_crit = sum(a['st'] == 'CRÍTICO' for a in achados)
        n_warn = sum(a['st'] == 'ATENÇÃO' for a in achados)
        n_ok = sum(a['st'] == 'OK' for a in achados)
        n_info = len(achados) - n_crit - n_warn - n_ok
        pct = 100 if n_ok + n_warn + n_crit == 0 else int((Decimal(100) * n_ok / (n_ok + n_warn + n_crit)).quantize(Decimal(1), ROUND_HALF_UP))

        # detalhes por guia
        det = OrderedDict()
        for idx, sec in enumerate(self.secoes, 1):
            g = guia_secao(sec['grupo'], sec['titulo'])
            det.setdefault(g, []).append(self.html_secao(idx, sec))

        cab = open(os.path.join(AQUI, 'template', 'cabecalho.html'), encoding='utf-8').read()
        h = '<!DOCTYPE html>\n' + cab.replace('{{TITULO}}', esc('Assessment SQL Server - ' + servidor))

        # capa
        versao_nome = self.dash('SQL Server') if self.dashboard is not None else None
        if not versao_nome or not versao_nome.startswith('SQL Server'):
            maj = (self.prop('ProductVersion') or '').split('.')[0]
            versao_nome = {'17': 'SQL Server 2025', '16': 'SQL Server 2022', '15': 'SQL Server 2019',
                           '14': 'SQL Server 2017', '13': 'SQL Server 2016'}.get(maj, 'SQL Server')
        upd = self.prop('ProductUpdateLevel')
        versao = f"{versao_nome} · {self.prop('ProductVersion') or ''} · {self.prop('ProductLevel') or ''}" + (f' {upd}' if upd else '')
        hi = self.rs('Sistema operacional')
        so = (hi.primeiro('host_distribution') if hi else None) or self.dash('Distribuição') or 'N/D'
        rel = (hi.primeiro('host_release') if hi else None) or self.dash('Release')
        if rel and rel != 'N/D':
            so += f' ({rel})'
        virt = self.hw('virtual_machine_type_desc') or self.dash('Virtualização') or 'NONE'
        aut = 'Somente Windows' if self.prop('IsIntegratedSecurityOnly') == '1' else 'Windows + SQL Server (misto)'
        h += ('<header class="cover"><div class="kicker">Assessment técnico · SQL Server</div>'
              '<h1>Levantamento do Ambiente de Banco de Dados</h1>'
              + (f'<div class="client">Cliente: <b>{esc(self.cliente)}</b></div>' if self.cliente.strip() else '')
              + f'<div class="server"><span class="lbl">Servidor \\ Instância</span><span class="val">{esc(servidor)}</span></div>'
              '<div class="meta">'
              f'<div><span>Data da coleta</span><b>{data_coleta}</b></div>'
              f'<div><span>Versão</span><b>{esc(versao)}</b></div>'
              f'<div><span>Edição</span><b>{esc(self.prop("Edition") or self.dash("Edition") or "N/D")}</b></div>'
              f'<div><span>Script de coleta</span><b>Assessment {esc(ver)} (saída .rpt)</b></div>'
              f'<div><span>Sistema operacional</span><b>{esc(so)}</b></div>'
              f'<div><span>Virtualização</span><b>{esc(virt)}</b></div>'
              f'<div><span>Collation</span><b>{esc(self.prop("Collation") or "N/D")}</b></div>'
              f'<div><span>Autenticação</span><b>{aut}</b></div>'
              '</div></header>')

        # guias
        nav = ''
        for pos, gid, rot, curto, lead in GUIAS:
            c = sum(a['st'] == 'CRÍTICO' and (a['guia'] == gid or gid == 'atencao') for a in achados)
            w = sum(a['st'] == 'ATENÇÃO' and (a['guia'] == gid or gid == 'atencao') for a in achados)
            badge = f'<span class="n {"crit" if c else "warn"}">{c + w}</span>' if gid != 'checklist' and c + w > 0 else ''
            nav += f'<button data-t="{gid}" onclick="tab(\'{gid}\')" type="button">{curto}{badge}</button>'
        h += '<nav class="tabs">' + nav + '</nav>'

        # 01 - resumo executivo
        cor = '#2E7D4F' if pct >= 80 else '#A86400' if pct >= 50 else '#B3121B'
        dash = (Decimal('326.73') * pct / 100).quantize(Decimal('0.1'), ROUND_HALF_UP)
        h += ('<div class="tabpane" id="t-geral"><section id="resumo"><h2><span class="num">01</span>Resumo executivo</h2>'
              '<div class="exec"><div class="gauge"><svg viewbox="0 0 140 140"><circle cx="70" cy="70" fill="none" r="52" stroke="#E4E4E1" stroke-width="14"></circle>'
              f'<circle cx="70" cy="70" fill="none" r="52" stroke="{cor}" stroke-dasharray="{dash} 326.73" stroke-linecap="round" stroke-width="14" transform="rotate(-90 70 70)"></circle>'
              f'<text class="pct" text-anchor="middle" x="70" y="80">{pct}%</text></svg><div class="cap">das verificações automáticas em conformidade</div></div><div class="counts">'
              f'<div class="cnt crit"><span>Crítico</span><b>{n_crit}</b><em>ação imediata</em></div>'
              f'<div class="cnt warn"><span>Atenção</span><b>{n_warn}</b><em>revisar</em></div>'
              f'<div class="cnt ok"><span>OK</span><b>{n_ok}</b><em>em conformidade</em></div>'
              f'<div class="cnt"><span>Info</span><b>{n_info}</b><em>informativo</em></div></div></div>')
        h += self.kpis(coleta)
        h += '</section>'
        linhas_g = ''.join(self.linha_achado(a) for a in achados if a['guia'] == 'geral')
        if linhas_g or det.get('geral'):
            h += '<section>'
            if linhas_g:
                h += '<div class="sub-h">Verificações desta área</div>' + CABEC_ACHADOS + linhas_g + '</tbody></table></div>'
            if det.get('geral'):
                h += '<div class="sub-h">Detalhamento técnico</div>' + TOOLS + ''.join(det['geral'])
            h += '</section>'
        h += '</div>'

        # 02 - pontos de atenção
        aten = sorted([a for a in achados if a['st'] in ('CRÍTICO', 'ATENÇÃO')],
                      key=lambda a: (0 if a['st'] == 'CRÍTICO' else 1, a['ordem']))
        h += ('<div class="tabpane" id="t-atencao"><section id="atencao"><h2><span class="num">02</span>Pontos de atenção</h2>'
              '<p class="lead">Itens classificados automaticamente como CRÍTICO ou ATENÇÃO. A classificação é indicativa e deve ser validada com o contexto do ambiente e do negócio.</p>')
        if aten:
            h += '<div class="tw"><table><thead><tr><th>Status</th><th>Categoria</th><th>Indicador</th><th>Valor</th><th>Observação</th></tr></thead><tbody>'
            for a in aten:
                h += (f'<tr><td><span class="b {"crit" if a["st"] == "CRÍTICO" else "warn"}">{esc(a["st"])}</span></td>'
                      f'<td>{esc(a["cat"])}</td><td>{esc(a["ind"])}</td><td class="n">{esc(a["val"]) or "—"}</td>'
                      f'<td class="obs">{esc(a["obs"]) or "—"}</td></tr>')
            h += '</tbody></table></div>'
        else:
            h += '<div class="note">Nenhum ponto de atenção identificado pelas verificações automáticas.</div>'
        h += '</section></div>'

        # 03 a 10
        for pos, gid, rot, curto, lead in GUIAS[2:10]:
            linhas = ''.join(self.linha_achado(a) for a in sorted(achados, key=lambda a: a['ordem']) if a['guia'] == gid)
            h += (f'<div class="tabpane" id="t-{gid}"><section><h2><span class="num">{pos:02d}</span>{rot}</h2>'
                  f'<p class="lead">{lead}</p>')
            if linhas:
                h += '<div class="sub-h">Verificações desta área</div>' + CABEC_ACHADOS + linhas + '</tbody></table></div>'
            if gid == 'armazenamento':
                h += '<div class="sub-h">Capacidade</div>' + self.capacidade()
            if gid == 'performance':
                h += '<div class="sub-h">Uso de CPU</div>' + self.cpu()
            if det.get(gid):
                h += '<div class="sub-h">Detalhamento técnico</div>' + TOOLS + ''.join(det[gid])
            h += '</section></div>'

        # 11 - checklist
        h += ('<div class="tabpane" id="t-checklist"><section id="checklist"><h2><span class="num">11</span>Checklist do assessment</h2>'
              '<p class="lead">Todas as verificações executadas, agrupadas por categoria.</p>' + CABEC_ACHADOS)
        if achados:
            ordem_cat = {}
            for a in achados:
                ordem_cat[a['cat']] = min(ordem_cat.get(a['cat'], a['ordem']), a['ordem'])
            vistos = set()
            for a in sorted(achados, key=lambda a: (ordem_cat[a['cat']], a['ordem'])):
                if a['cat'] not in vistos:
                    vistos.add(a['cat'])
                    h += f'<tr class="grp"><td colspan="4">{esc(a["cat"])}</td></tr>'
                h += self.linha_achado(a)
        else:
            h += '<tr><td class="empty" colspan="4">Sem dados.</td></tr>'
        h += '</tbody></table></div></section>'

        # 12 - observações da coleta
        obs = [('Origem dos dados', f'Arquivo {os.path.basename(self.nome_arquivo)} gerado pelo script {ver}. '
                f'Relatório montado a partir de {len(self.resultsets)} resultsets do .rpt; '
                'a lista completa de cada seção está no arquivo original.')]
        if self.summary is None:
            obs.append(('Checklist', 'O resultset do checklist (#AssessmentSummary) não foi encontrado no .rpt; '
                        'os indicadores foram calculados a partir do dashboard.'))
        for rs in self.nao_mapeados:
            obs.append(('Resultset não mapeado', f'Colunas: {", ".join(rs.colunas[:8])} (linha {rs.linha_arquivo} do .rpt). '
                        'Exibido na Visão Geral; atualizar o catálogo do gerador.'))
        for k, m in enumerate(self.mensagens):
            if re.match(r'^Msg \d+, (Level|Nível) \d+', m):
                det_msg = self.mensagens[k + 1] if k + 1 < len(self.mensagens) else ''
                obs.append((m, det_msg))
        h += ('<section id="coleta"><h2><span class="num">12</span>Observações da coleta</h2>'
              '<p class="lead">Informações sobre a origem dos dados e itens que não puderam ser coletados.</p>'
              '<div class="tw"><table><thead><tr><th>Seção</th><th>Mensagem</th></tr></thead><tbody>'
              + ''.join(f'<tr><td>{esc(a)}</td><td class="obs">{esc(b)}</td></tr>' for a, b in obs)
              + '</tbody></table></div></section></div>')

        rodape = open(os.path.join(AQUI, 'template', 'rodape_solo.b64'), encoding='ascii').read().strip()
        h += ('<footer class="foot"><img alt="Solo Network - Rod. João Leopoldo Jacomel, 12.162 - Pinhais/PR - 0800 604 9596 - www.solonetwork.com.br" '
              f'src="data:image/png;base64,{rodape}"/><div class="gen">SQL Server Assessment {esc(ver)} · dados coletados em {data_coleta} · '
              'coleta somente leitura</div></footer></div></body></html>')
        return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', ' ', h)

    def kpis(self, coleta):
        def g(v, f=str):
            try:
                return f(v) if v not in (None, '') else None
            except (ValueError, InvalidOperation):
                return None

        # log shipping só é coletado a partir do v11 (grupo 25)
        self.tem_log_shipping = any(sec['grupo'] == '25' for sec in self.secoes)
        cpus = self.hw('LogicalCPUCount') or self.dash('CPU') or '-'
        sub_cpu = f"{self.hw('socket_count') or '?'} socket(s) · {self.hw('cores_per_socket') or '?'} cores/socket · {self.hw('numa_node_count') or '?'} NUMA"
        mem_mb = g(self.hw('PhysicalMemoryMB'), Decimal)
        mem = fmt_num(mem_mb / 1024, 1) if mem_mb is not None else '-'
        cfg = self.rs('Configurações relevantes (sp_configure)')
        maxmem = None
        if cfg:
            for r in cfg.linhas:
                if (r[cfg.col('name')] or '').lower() == 'max server memory (mb)':
                    maxmem = g(r[cfg.col('RunningValue')], Decimal)
        if maxmem is None:
            v = self.dash('max server memory (MB)')
            maxmem = g(v, Decimal)
        sub_mem = ('max server memory: não configurado' if maxmem is not None and maxmem >= 2147483647
                   else 'max server memory: ' + (fmt_num(maxmem / 1024, 1) if maxmem is not None else '?') + ' GB')
        dbs = self.rs('Databases')
        n_db = self.dash('Databases') if self.dashboard is not None else None
        fora = 0
        if dbs:
            ids = [(g(r[dbs.col('database_id')], int) or 0, r[dbs.col('state_desc')]) for r in dbs.linhas]
            if n_db is None:
                n_db = str(sum(1 for i, s in ids if i > 4))
            fora = sum(1 for i, s in ids if i > 4 and s != 'ONLINE')
        arq = self.rs('Arquivos de dados e log')
        dados = log = None
        if arq:
            dados = log = Decimal(0)
            for r in arq.linhas:
                mb = g(r[arq.col('SizeMB')], Decimal) or Decimal(0)
                if r[arq.col('type_desc')] == 'LOG':
                    log += mb
                elif r[arq.col('type_desc')] == 'ROWS':
                    dados += mb
            dados = (dados / 1024).quantize(Decimal('0.01'), ROUND_HALF_UP)
            log = (log / 1024).quantize(Decimal('0.01'), ROUND_HALF_UP)
        hadr = (self.prop('IsHadrEnabled') == '1') or (self.dash('Always On / HADR') == 'HABILITADO')
        ags = self.dash('Availability Groups') or str(len(self.rs('Availability Groups').linhas) if self.rs('Availability Groups') else 0)
        reps = self.dash('Réplicas') or str(len(self.rs('Réplicas dos AGs').linhas) if self.rs('Réplicas dos AGs') else 0)
        fci = self.prop('IsClustered') == '1'
        lsp = len(self.rs('Log shipping - databases primárias').linhas) if self.rs('Log shipping - databases primárias') else 0
        lss = len(self.rs('Log shipping - databases secundárias').linhas) if self.rs('Log shipping - databases secundárias') else 0
        jobs = self.rs('Jobs do SQL Agent')
        n_jobs = self.dash('SQL Agent Jobs') or self.dash('Jobs totais') or (str(len(jobs.linhas)) if jobs else None)
        falha = rs_srv = 0
        if jobs:
            ie, ist, ic = jobs.col('JobEnabled'), jobs.col('LastRunStatus'), jobs.col('CategoryName')
            falha = sum(1 for r in jobs.linhas if r[ie] == '1' and r[ist] == '0') if ie >= 0 and ist >= 0 else 0
            rs_srv = sum(1 for r in jobs.linhas if r[ic] == 'Report Server') if ic >= 0 else 0
        sub_jobs = (f'{falha} com falha na última execução' if falha else 'nenhum com falha na última execução')
        if rs_srv:
            sub_jobs += f' · {fmt_int_milhar(rs_srv)} do Report Server'
        linked = self.dash('Linked Servers') or self.dash('Quantidade') or '-'
        inicio = parse_dt(self.hw('sqlserver_start_time') or '')
        uptime = f'{(coleta.date() - inicio.date()).days} dias' if inicio else '- dias'
        desde = fmt_dt(inicio, False) if inicio else '-'

        def kpi(k, v, s):
            return f'<div class="kpi"><span class="k">{k}</span><span class="v">{esc(v)}</span><span class="s">{esc(s)}</span></div>'

        return ('<div class="kpis">'
                + kpi('CPUs lógicas', cpus, sub_cpu)
                + kpi('Memória física', mem + ' GB', sub_mem)
                + kpi('Databases de usuário', n_db or '-', f'{fora} fora de ONLINE' if fora else 'todas ONLINE')
                + kpi('Tamanho alocado', (fmt_num(dados + log, 1) if dados is not None else '-') + ' GB',
                      f'dados {fmt_num(dados, 1) if dados is not None else "-"} GB · log {fmt_num(log, 1) if log is not None else "-"} GB')
                + kpi('Always On', 'Habilitado' if hadr else 'Não habilitado',
                      f'{ags} AG(s) · {reps} réplica(s)' + (' · instância FCI' if fci else '')
                      + (f' · log shipping: {lsp} primária(s), {lss} secundária(s)' if lsp + lss
                         else ' · sem log shipping' if self.tem_log_shipping else ''))
                + kpi('Jobs do SQL Agent', fmt_int_milhar(n_jobs) if n_jobs and RE_INT.match(n_jobs) else (n_jobs or '-'), sub_jobs)
                + kpi('Linked servers', linked, 'não sistêmicos')
                + kpi('Uptime', uptime, 'desde ' + desde)
                + '</div>')


def fmt_ponto(d):
    return str(Decimal(d).quantize(Decimal('0.1'), ROUND_HALF_UP))


def main():
    ap = argparse.ArgumentParser(description='Gera o HTML do SQL Server Assessment (Solo) a partir do .rpt do SSMS.')
    ap.add_argument('rpt', help='arquivo .rpt (Resultados em arquivo do SSMS)')
    ap.add_argument('--cliente', default='', help='nome do cliente exibido na capa')
    ap.add_argument('--saida', help='arquivo .html de saída (padrão: Assessment_<cliente>_<instância>.html)')
    a = ap.parse_args()
    rel = Relatorio(a.rpt, a.cliente, a.rpt)
    if not rel.secoes and rel.html_pronto.rstrip().endswith('</html>'):
        # .rpt gerado com @SomenteHTML = 1 e o HTML veio inteiro: usa o próprio relatório do script
        html = '<!DOCTYPE html>\n' + rel.html_pronto
    elif not rel.secoes and rel.summary is None:
        sys.exit('Nenhum resultset do assessment encontrado no .rpt (o HTML do script, se houver, está truncado). '
                 'Reexecute com @SomenteHTML = 0 e saída "Resultados em arquivo".')
    else:
        html = rel.gerar()
    if not a.saida:
        inst = re.sub(r'[\\/:*?"<>|, ]', '_', rel.prop('ServerName') or 'instancia')
        cli = re.sub(r'[\\/:*?"<>|, ]', '_', a.cliente) + '_' if a.cliente else ''
        a.saida = os.path.join(os.path.dirname(os.path.abspath(a.rpt)), f'Assessment_{cli}{inst}.html')
    with open(a.saida, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Relatório gerado: {a.saida}')
    print(f'  resultsets lidos: {len(rel.resultsets)} | seções: {len(rel.secoes)} | '
          f'não mapeados: {len(rel.nao_mapeados)} | checklist: {"sim" if rel.summary else "não"}')
    for rs in rel.nao_mapeados:
        print(f'  AVISO: resultset não mapeado na linha {rs.linha_arquivo}: {rs.colunas[:5]}', file=sys.stderr)


if __name__ == '__main__':
    main()
