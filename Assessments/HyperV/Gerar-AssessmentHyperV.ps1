<#
.SYNOPSIS
    Hyper-V Assessment v1 (Solo Network): coleta somente leitura do host Hyper-V e gera o relatório HTML.

.DESCRIPTION
    Executa no próprio host Hyper-V (Windows Server 2016/2019/2022) e levanta:
      - Host: hardware, sistema operacional, atualizações, reinício pendente, plano de energia, hora, antivírus, serviços.
      - Hyper-V: configuração do host, versão padrão de configuração, recursos instalados, serviço VMMS.
      - Máquinas virtuais: estado, heartbeat, versão, vCPU, memória, integration services, sistema convidado,
        ações automáticas, tipo de checkpoint, firmware, DVD/ISO.
      - Armazenamento: volumes e espaço livre, discos físicos, VHD/VHDX (tipo, tamanho, fragmentação, cadeia),
        sobrecarga de discos dinâmicos, arquivos de disco não anexados.
      - Rede: switches virtuais, adaptadores físicos, VMQ, teaming, IPs, adaptadores e VLANs das VMs.
      - Checkpoints, Hyper-V Replica, VSS writers e software de backup.
      - Eventos (Hyper-V, System e Application) e uma amostra de desempenho (CPU do hipervisor, latência de disco).

    Somente leitura: o script não altera nenhuma configuração do servidor nem das VMs.
    Usa apenas módulos nativos do Windows Server (Hyper-V, Storage, NetAdapter, CIM).
    Requer PowerShell executado como Administrador.

.EXAMPLE
    .\Gerar-AssessmentHyperV.ps1 -Cliente "Cliente X" -Abrir

.EXAMPLE
    .\Gerar-AssessmentHyperV.ps1 -Cliente "Cliente X" -PastaSaida C:\Temp -DiasEventos 15 -AmostraSegundos 30
#>
[CmdletBinding()]
param(
    [string]$Cliente = '',

    [string]$PastaSaida = (Get-Location).Path,

    # Janela (em dias) para a análise dos logs de eventos.
    [int]$DiasEventos = 30,

    # Duração (em segundos) da amostra de desempenho (CPU do hipervisor e latência de disco).
    [int]$AmostraSegundos = 15,

    # Abre o relatório no navegador ao final.
    [switch]$Abrir
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

$Script:Versao  = 'Hyper-V Assessment v1'
$Script:CI      = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
$Script:Agora   = Get-Date
$Script:Checks  = New-Object System.Collections.Generic.List[object]
$Script:Notas   = New-Object System.Collections.Generic.List[object]
$Script:SecN    = 0
$Script:Volumes = @()

# =================================================================================================
#  Funções auxiliares
# =================================================================================================
function Esc([object]$s) { if ($null -eq $s) { return '' }; [System.Net.WebUtility]::HtmlEncode([string]$s) }

function Num($v, [int]$d = 0) {
    if ($null -eq $v -or "$v" -eq '') { return '' }
    try { ([double]$v).ToString('N' + $d, $Script:CI) } catch { [string]$v }
}

function Bytes($b, [int]$d = 1) {
    if ($null -eq $b -or "$b" -eq '') { return '' }
    $b = [double]$b
    if ($b -ge 1TB) { return (Num ($b / 1TB) 2) + ' TB' }
    if ($b -ge 1GB) { return (Num ($b / 1GB) $d) + ' GB' }
    if ($b -ge 1MB) { return (Num ($b / 1MB) 0) + ' MB' }
    return (Num ($b / 1KB) 0) + ' KB'
}

function Dt($d) {
    if (-not $d) { return '' }
    try { ([datetime]$d).ToString('dd/MM/yyyy HH:mm') } catch { [string]$d }
}

function Span($ts) {
    if ($null -eq $ts) { return '' }
    $ts = [timespan]$ts
    if ($ts.TotalSeconds -le 0) { return '' }
    if ($ts.Days -gt 0) { return ('{0} d {1} h' -f $ts.Days, $ts.Hours) }
    return ('{0} h {1} min' -f $ts.Hours, $ts.Minutes)
}

function SN($b) { if ($b) { 'SIM' } else { 'NÃO' } }

function Pct($parte, $total) { if (-not $total) { return 0 }; [math]::Round(100.0 * $parte / $total, 1) }

# Primeiro valor não nulo entre as propriedades informadas.
function PV($o, [string[]]$nomes) {
    if ($null -eq $o) { return $null }
    foreach ($n in $nomes) {
        $p = $o.PSObject.Properties[$n]
        if ($p -and $null -ne $p.Value -and "$($p.Value)" -ne '') { return $p.Value }
    }
    return $null
}

function Nota([string]$secao, [string]$msg) {
    $Script:Notas.Add([pscustomobject]@{ 'Seção' = $secao; 'Mensagem' = $msg })
}

# Executa um bloco de coleta; qualquer erro vira observação da coleta e o relatório continua.
function Coletar([string]$secao, [scriptblock]$bloco) {
    $ErrorActionPreference = 'Stop'
    try { return (& $bloco) }
    catch { Nota $secao $_.Exception.Message; return $null }
}

function Passo([string]$t) { Write-Host ('  - ' + $t) -ForegroundColor DarkGray }

# Area = guia do relatório; Status = CRIT | WARN | OK | INFO
function Add-Check([string]$Area, [string]$Cat, [string]$Ind, $Val, [string]$Status, [string]$Obs) {
    $Script:Checks.Add([pscustomobject]@{ Area = $Area; Cat = $Cat; Ind = $Ind; Val = [string]$Val; Status = $Status; Obs = $Obs })
}

function Badge([string]$st, [string]$txt) {
    $cls = switch ($st) { 'CRIT' { 'crit' } 'WARN' { 'warn' } 'OK' { 'ok' } default { 'info' } }
    if (-not $txt) { $txt = switch ($st) { 'CRIT' { 'CRÍTICO' } 'WARN' { 'ATENÇÃO' } 'OK' { 'OK' } default { 'INFO' } } }
    '<span class="b ' + $cls + '">' + (Esc $txt) + '</span>'
}

# Tabela genérica a partir de objetos. $Num = colunas numéricas, $Raw = colunas com HTML pronto, $Wide = texto longo.
function Tbl($rows, [string[]]$Cols, [string[]]$Num = @(), [string[]]$Raw = @(), [string[]]$Wide = @(), [switch]$Data) {
    $rows = @($rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return '<p class="empty">Nenhum registro.</p>' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tw"><table' + $(if ($Data) { ' class="data"' } else { '' }) + '><thead><tr>')
    foreach ($c in $Cols) { [void]$sb.Append('<th>' + (Esc $c) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $rows) {
        [void]$sb.Append('<tr>')
        foreach ($c in $Cols) {
            $v = $r.$c
            $cls = @()
            if ($Num -contains $c) { $cls += 'n' }
            if ($Wide -contains $c) { $cls += 'w' }
            $td = if ($cls.Count) { '<td class="' + ($cls -join ' ') + '">' } else { '<td>' }
            if ($Raw -contains $c) { [void]$sb.Append($td + [string]$v + '</td>') }
            else { [void]$sb.Append($td + (Esc $v) + '</td>') }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    $sb.ToString()
}

# Tabela vertical propriedade/valor.
function KV($pares) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tw"><table class="data kv"><thead><tr><th>Propriedade</th><th>Valor</th></tr></thead><tbody>')
    foreach ($k in $pares.Keys) { [void]$sb.Append('<tr><th>' + (Esc $k) + '</th><td>' + (Esc $pares[$k]) + '</td></tr>') }
    [void]$sb.Append('</tbody></table></div>')
    $sb.ToString()
}

function Det([string]$num, [string]$titulo, [string]$desc, [string]$html, $qtd) {
    $Script:SecN++
    $c = if ($null -eq $qtd) { '' } elseif ($qtd -is [string]) { $qtd } elseif ($qtd -eq 1) { '1 registro' } else { "$qtd registros" }
    $d = if ($desc) { '<p class="desc">' + (Esc $desc) + '</p>' } else { '' }
    '<details id="s' + $Script:SecN + '" open="open"><summary><span class="g">' + $num + '</span>' + (Esc $titulo) +
    '<span class="c">' + (Esc $c) + '</span></summary>' + $d + $html + '</details>'
}

# Barras horizontais: itens com N (nome), P (0-100), C (classe ok/warn/crit/db) e T (texto à direita).
function Bars($itens) {
    $itens = @($itens)
    if ($itens.Count -eq 0) { return '<p class="empty">Sem dados.</p>' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="hb">')
    foreach ($i in $itens) {
        $p = [math]::Max([double]0, [math]::Min([double]100, [math]::Round([double]$i.P)))
        [void]$sb.Append('<span class="n" title="' + (Esc $i.N) + '">' + (Esc $i.N) + '</span><div class="bar"><i class="' + $i.C +
            '" style="width:' + $p + '%">' + $p + '</i></div><span class="t">' + (Esc $i.T) + '</span>')
    }
    [void]$sb.Append('</div>')
    $sb.ToString()
}

function Card([string]$titulo, [string]$html) { '<div class="card"><h3>' + (Esc $titulo) + '</h3>' + $html + '</div>' }

function Kpi([string]$k, [string]$v, [string]$s) {
    '<div class="kpi"><span class="k">' + (Esc $k) + '</span><span class="v">' + (Esc $v) + '</span><span class="s">' + (Esc $s) + '</span></div>'
}

function ChecksArea([string]$area) {
    $c = @($Script:Checks | Where-Object { $_.Area -eq $area })
    if ($c.Count -eq 0) { return '' }
    $rows = foreach ($x in $c) {
        [pscustomobject]@{ 'Indicador' = $x.Ind; 'Valor' = $x.Val; 'Status' = (Badge $x.Status); 'Observação' = $x.Obs }
    }
    '<div class="sub-h">Verificações desta área</div>' + (Tbl $rows 'Indicador', 'Valor', 'Status', 'Observação' -Num 'Valor' -Raw 'Status')
}

function Ferramentas { '<div class="sub-h">Detalhamento técnico</div><div class="tools"><button onclick="tg(true)" type="button">Expandir tudo</button><button onclick="tg(false)" type="button">Recolher tudo</button></div>' }

function VolumeDe([string]$caminho) {
    if (-not $caminho) { return $null }
    $melhor = $null
    foreach ($v in $Script:Volumes) {
        if ($v.Nome -and $v.Nome -notlike '\\?\*' -and $caminho.StartsWith($v.Nome, [StringComparison]::OrdinalIgnoreCase)) {
            if (-not $melhor -or $v.Nome.Length -gt $melhor.Nome.Length) { $melhor = $v }
        }
    }
    $melhor
}

function Get-Eventos([string]$log, [int[]]$niveis, [datetime]$desde, [int]$max = 3000) {
    try {
        @(Get-WinEvent -FilterHashtable @{ LogName = $log; Level = $niveis; StartTime = $desde } -MaxEvents $max -ErrorAction Stop)
    } catch {
        if ($_.FullyQualifiedErrorId -notmatch 'NoMatchingEventsFound' -and $_.Exception.Message -notmatch 'No events were found|Nenhum evento') {
            Nota "Eventos - $log" $_.Exception.Message
        }
        @()
    }
}

function NivelEvento($l) { switch ([int]$l) { 1 { 'Crítico' } 2 { 'Erro' } 3 { 'Aviso' } default { "Nível $l" } } }

function Resumir-Eventos($eventos) {
    $eventos = @($eventos)
    if ($eventos.Count -eq 0) { return @() }
    $eventos | Group-Object LogName, ProviderName, Id, Level | ForEach-Object {
        $g = @($_.Group | Sort-Object TimeCreated)
        $ult = $g[-1]
        $msg = [string]$ult.Message
        if (-not $msg) { $msg = '(mensagem não disponível)' }
        $msg = ($msg -split "`r?`n")[0]
        if ($msg.Length -gt 320) { $msg = $msg.Substring(0, 320) + '…' }
        $st = switch ([int]$ult.Level) { 1 { 'CRIT' } 2 { 'CRIT' } default { 'WARN' } }
        [pscustomobject]@{
            'Log' = $ult.LogName; 'Origem' = $ult.ProviderName; 'ID' = $ult.Id
            'Nível' = (Badge $st (NivelEvento $ult.Level)); 'Ocorrências' = $g.Count
            'Primeira' = (Dt $g[0].TimeCreated); 'Última' = (Dt $ult.TimeCreated); 'Mensagem (última)' = $msg
            '_n' = $g.Count; '_lv' = [int]$ult.Level
        }
    } | Sort-Object _lv, @{ Expression = '_n'; Descending = $true }
}

function Get-GuestKvp($vm) {
    $r = @{}
    $cs = Get-CimInstance -Namespace root\virtualization\v2 -ClassName Msvm_ComputerSystem -Filter ("Name='{0}'" -f $vm.Id) -ErrorAction Stop
    if (-not $cs) { return $r }
    $kvp = Get-CimAssociatedInstance -InputObject $cs -ResultClassName Msvm_KvpExchangeComponent -ErrorAction Stop
    foreach ($item in @($kvp.GuestIntrinsicExchangeItems)) {
        if (-not $item) { continue }
        $x = [xml]$item
        $n = ($x.INSTANCE.PROPERTY | Where-Object { $_.NAME -eq 'Name' }).VALUE
        $d = ($x.INSTANCE.PROPERTY | Where-Object { $_.NAME -eq 'Data' }).VALUE
        if ($n) { $r[$n] = $d }
    }
    $r
}

# Fim do suporte estendido do Windows Server (Microsoft Lifecycle).
function Ciclo-Vida([string]$caption) {
    $tab = [ordered]@{ '2025' = '2034-11-14'; '2022' = '2031-10-14'; '2019' = '2029-01-09'; '2016' = '2027-01-12'; '2012' = '2023-10-10'; '2008' = '2020-01-14' }
    foreach ($k in $tab.Keys) { if ($caption -match $k) { return [datetime]$tab[$k] } }
    $null
}

# =================================================================================================
#  Coleta e geração do relatório
# =================================================================================================
function Main {
    # --- Pré-requisitos -------------------------------------------------------------------------
    $principal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Execute o PowerShell como Administrador (botão direito > Executar como administrador).'
    }
    if (-not (Get-Command Get-VM -ErrorAction SilentlyContinue)) {
        Import-Module Hyper-V -ErrorAction Stop
    }
    if (-not (Test-Path -LiteralPath $PastaSaida)) { New-Item -ItemType Directory -Path $PastaSaida | Out-Null }

    $desde = $Script:Agora.AddDays(-$DiasEventos)
    $sysDrive = ($env:SystemDrive + '\')
    Write-Host "$($Script:Versao) - coleta somente leitura" -ForegroundColor Cyan

    # ============================================================================================
    #  HOST
    # ============================================================================================
    Passo 'Host e sistema operacional'
    $cs   = Coletar 'Host - Win32_ComputerSystem'   { Get-CimInstance Win32_ComputerSystem }
    $os   = Coletar 'Host - Win32_OperatingSystem'  { Get-CimInstance Win32_OperatingSystem }
    $bios = Coletar 'Host - Win32_BIOS'             { Get-CimInstance Win32_BIOS }
    $cpus = @(Coletar 'Host - Win32_Processor'      { Get-CimInstance Win32_Processor })
    $dimm = @(Coletar 'Host - Win32_PhysicalMemory' { Get-CimInstance Win32_PhysicalMemory })
    $nt   = Coletar 'Host - registro CurrentVersion' { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' }
    $pf   = @(Coletar 'Host - Win32_PageFileUsage'  { Get-CimInstance Win32_PageFileUsage })

    $hostNome = if ($cs) { $cs.Name } else { $env:COMPUTERNAME }
    $fqdn = if ($cs -and $cs.PartOfDomain) { "$($cs.Name).$($cs.Domain)" } else { $hostNome }
    $memTotal = if ($cs) { [double]$cs.TotalPhysicalMemory } else { 0.0 }
    $memLivre = if ($os) { [double]$os.FreePhysicalMemory * 1KB } else { 0.0 }
    $lp      = ($cpus | Measure-Object NumberOfLogicalProcessors -Sum).Sum
    $cores   = ($cpus | Measure-Object NumberOfCores -Sum).Sum
    $sockets = $cpus.Count
    $build   = if ($nt) { "$($nt.CurrentBuildNumber).$($nt.UBR)" } elseif ($os) { $os.BuildNumber } else { '' }
    $release = if ($nt) { PV $nt 'DisplayVersion', 'ReleaseId' } else { '' }
    $osNome  = if ($os) { $os.Caption } else { '' }
    $boot    = if ($os) { $os.LastBootUpTime } else { $null }
    $uptime  = if ($boot) { $Script:Agora - $boot } else { $null }

    $Area = 'geral'; $C1 = '01 - Host'; $C2 = '02 - Sistema operacional'; $C3 = '03 - Hyper-V'
    Add-Check $Area $C1 'Servidor' $fqdn 'INFO' 'Nome do host Hyper-V.'
    Add-Check $Area $C1 'Fabricante / modelo' ("$(PV $cs 'Manufacturer') $(PV $cs 'Model')").Trim() 'INFO' 'Validar firmware/BIOS e drivers com o fabricante.'
    Add-Check $Area $C1 'Processadores' ("{0} socket(s) · {1} cores · {2} lógicos" -f $sockets, $cores, $lp) 'INFO' $(if ($cpus) { $cpus[0].Name.Trim() } else { '' })
    $ht = ($lp -gt $cores)
    Add-Check $Area $C1 'Hyper-Threading' (SN $ht) 'INFO' 'Processadores lógicos acima do número de cores indicam Hyper-Threading ativo.'
    $pctLivre = Pct $memLivre $memTotal
    $stMem = if ($pctLivre -lt 5 -or $memLivre -lt 1GB) { 'CRIT' } elseif ($pctLivre -lt 10 -or $memLivre -lt 2GB) { 'WARN' } else { 'OK' }
    Add-Check $Area $C1 'Memória livre no host' ("{0} de {1} ({2}%)" -f (Bytes $memLivre), (Bytes $memTotal), (Num $pctLivre 1)) $stMem 'O host (partição raiz) precisa de memória livre para o VMMS, backup e para iniciar VMs. Abaixo de 10% merece revisão.'

    Add-Check $Area $C2 'Sistema operacional' ("$osNome · build $build" + $(if ($release) { " · $release" } else { '' })) 'INFO' 'Versão e build (UBR) indicam o nível de atualização cumulativa.'
    $fim = Ciclo-Vida $osNome
    if ($fim) {
        $meses = (($fim - $Script:Agora).TotalDays / 30.4)
        $st = if ($meses -le 0) { 'CRIT' } elseif ($meses -lt 18) { 'WARN' } else { 'OK' }
        $obs = if ($meses -le 0) { 'Sistema operacional fora do suporte estendido da Microsoft; sem atualizações de segurança.' } elseif ($meses -lt 18) { 'Menos de 18 meses de suporte estendido; planejar migração.' } else { 'Dentro do suporte estendido da Microsoft.' }
        Add-Check $Area $C2 'Ciclo de vida (fim do suporte estendido)' $fim.ToString('dd/MM/yyyy') $st $obs
    }
    if ($uptime) {
        $st = if ($uptime.TotalDays -gt 90) { 'WARN' } else { 'OK' }
        Add-Check $Area $C2 'Uptime do host (dias)' ([math]::Floor($uptime.TotalDays)) $st 'Acima de 90 dias sem reinício normalmente indica atualizações acumuladas pendentes de aplicação.'
    }

    $hf = @(Coletar 'Host - Get-HotFix' { Get-HotFix } | Where-Object { $_ })
    $hfData = @($hf | Where-Object { $_.InstalledOn } | Sort-Object InstalledOn -Descending)
    if ($hfData.Count) {
        $dias = [math]::Floor(($Script:Agora - [datetime]$hfData[0].InstalledOn).TotalDays)
        $st = if ($dias -gt 120) { 'CRIT' } elseif ($dias -gt 60) { 'WARN' } else { 'OK' }
        Add-Check $Area $C2 'Última atualização instalada' ("{0} ({1} · há {2} dias)" -f (Dt $hfData[0].InstalledOn).Substring(0, 10), $hfData[0].HotFixID, $dias) $st 'Atualizações cumulativas são mensais. Acima de 60 dias o host está desatualizado.'
    } else {
        Add-Check $Area $C2 'Última atualização instalada' 'não identificada' 'WARN' 'Get-HotFix não retornou datas de instalação.'
    }

    $motivos = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $motivos += 'Component Based Servicing' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $motivos += 'Windows Update' }
    $pfro = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
    $st = if ($motivos.Count) { 'WARN' } else { 'OK' }
    if ($pfro) { $motivos += 'PendingFileRenameOperations' }
    Add-Check $Area $C2 'Reinício pendente' $(if ($motivos.Count) { 'SIM: ' + ($motivos -join ', ') } else { 'NÃO' }) $st 'Atualizações instaladas só entram em vigor após o reinício. Planejar janela (as VMs serão afetadas).'

    $plano = Coletar 'Host - plano de energia' { (& powercfg.exe /getactivescheme) -join ' ' }
    if ($plano) {
        $guid = if ($plano -match '([0-9a-fA-F]{8}-[0-9a-fA-F-]{27})') { $Matches[1].ToLower() } else { '' }
        $nomePlano = if ($plano -match '\(([^)]+)\)\s*$') { $Matches[1] } else { $plano }
        $st = if ($guid -in '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c', 'e9a42b02-d5df-448d-aa00-03f14749eb61') { 'OK' } else { 'WARN' }
        Add-Check $Area $C2 'Plano de energia' $nomePlano $st 'Para hosts de virtualização recomenda-se "Alto desempenho" (no Windows e no BIOS/UEFI) para evitar latência por economia de energia.'
    }

    $fonteHora = Coletar 'Host - w32tm' { ((& w32tm.exe /query /source) | Where-Object { $_ -and $_.Trim() }) -join ' ' }
    if ($fonteHora) {
        $st = if ($fonteHora -match 'CMOS|Free-running|Relógio') { 'WARN' } else { 'OK' }
        Add-Check $Area $C2 'Fonte de hora (w32tm)' $fonteHora.Trim() $st 'O host deve sincronizar com o domínio ou NTP confiável; VMs herdam a hora pelo integration service de sincronização.'
    }

    $mp = Coletar 'Host - Microsoft Defender' { Get-MpComputerStatus }
    $mpPref = if ($mp) { Coletar 'Host - Get-MpPreference' { Get-MpPreference } } else { $null }
    if ($mp) {
        $idade = if ($mp.AntivirusSignatureLastUpdated) { [math]::Floor(($Script:Agora - $mp.AntivirusSignatureLastUpdated).TotalDays) } else { $null }
        $st = if (-not $mp.RealTimeProtectionEnabled) { 'WARN' } elseif ($null -ne $idade -and $idade -gt 7) { 'WARN' } else { 'OK' }
        Add-Check $Area $C2 'Antivírus (Microsoft Defender)' ("Tempo real: {0} · assinaturas de {1}" -f (SN $mp.RealTimeProtectionEnabled), (Dt $mp.AntivirusSignatureLastUpdated)) $st 'Proteção em tempo real ativa e assinaturas com até 7 dias.'
        if ($mpPref -and $mpPref.DisableAutoExclusions) {
            Add-Check $Area $C2 'Exclusões automáticas do Defender' 'DESATIVADAS' 'WARN' 'Sem exclusões para VHDX, VMs e processos do Hyper-V (vmms.exe, vmwp.exe) o antivírus pode corromper ou travar VMs.'
        }
    } else {
        Add-Check $Area $C2 'Antivírus (Microsoft Defender)' 'não disponível' 'INFO' 'Provável antivírus de terceiros. Confirmar exclusões recomendadas para Hyper-V (VHD/VHDX/AVHDX, pastas das VMs, vmms.exe, vmwp.exe, vmcompute.exe).'
    }
    $fw = @(Coletar 'Host - firewall' { Get-NetFirewallProfile })

    # --- Hyper-V -------------------------------------------------------------------------------
    Passo 'Configuração do Hyper-V'
    $vmh = Coletar 'Hyper-V - Get-VMHost' { Get-VMHost }
    $verPadrao = Coletar 'Hyper-V - versão padrão de configuração' { (Get-VMHostSupportedVersion -Default).Version }
    $features = @(Coletar 'Host - recursos do Windows' {
        Get-WindowsFeature -Name Hyper-V, Hyper-V-PowerShell, RSAT-Hyper-V-Tools, Hyper-V-Tools, Failover-Clustering, Multipath-IO, Windows-Server-Backup, Data-Center-Bridging, FS-Data-Deduplication
    })
    $svcAll = @(Coletar 'Host - serviços' { Get-CimInstance Win32_Service })
    $vmms = $svcAll | Where-Object { $_.Name -eq 'vmms' }
    $st = if ($vmms -and $vmms.State -eq 'Running') { 'OK' } else { 'CRIT' }
    Add-Check $Area $C3 'Serviço Hyper-V Virtual Machine Management (vmms)' $(if ($vmms) { $vmms.State } else { 'não encontrado' }) $st 'Serviço que gerencia as VMs; deve estar em execução e com início automático.'
    $fHv = $features | Where-Object { $_.Name -eq 'Hyper-V' }
    if ($fHv) { Add-Check $Area $C3 'Função Hyper-V' $fHv.InstallState 'INFO' 'Estado da função Hyper-V no Windows Server.' }
    $clus = $features | Where-Object { $_.Name -eq 'Failover-Clustering' -and $_.Installed }
    Add-Check $Area $C3 'Failover Cluster' $(if ($clus) { 'Instalado' } else { 'Não instalado' }) 'INFO' $(if ($clus) { 'Host participa (ou pode participar) de cluster; validar Cluster-Aware Updating e CSV.' } else { 'Host standalone: uma falha do servidor para todas as VMs. Backup e/ou Hyper-V Replica são essenciais.' })
    if ($vmh) {
        Add-Check $Area $C3 'Versão padrão de configuração de VM' $verPadrao 'INFO' 'Versão mais alta suportada por este host; VMs com versão inferior não usam recursos novos.'
        foreach ($par in @(@('Pasta padrão de VHDs', $vmh.VirtualHardDiskPath), @('Pasta padrão de VMs', $vmh.VirtualMachinePath))) {
            $st = if ($par[1] -and $par[1].StartsWith($sysDrive, [StringComparison]::OrdinalIgnoreCase)) { 'WARN' } else { 'OK' }
            Add-Check $Area $C3 $par[0] $par[1] $st 'Pasta padrão no volume do sistema faz novas VMs e discos serem criados em C:. Recomenda-se um volume dedicado.'
        }
        Add-Check $Area $C3 'NUMA spanning' (SN $vmh.NumaSpanningEnabled) 'INFO' 'Permite que VMs usem memória de outro nó NUMA (mais flexível, porém com possível perda de desempenho).'
        Add-Check $Area $C3 'Live Migration' $(if ($vmh.VirtualMachineMigrationEnabled) { "Habilitada ($($vmh.VirtualMachineMigrationAuthenticationType))" } else { 'Desabilitada' }) 'INFO' 'Migração de VMs entre hosts.'
    }

    # ============================================================================================
    #  ARMAZENAMENTO (volumes primeiro: usados no mapeamento de VHDs)
    # ============================================================================================
    Passo 'Volumes e discos'
    $w32vol = @(Coletar 'Armazenamento - Win32_Volume' { Get-CimInstance Win32_Volume -Filter 'DriveType=3' })
    $gvol   = @(Coletar 'Armazenamento - Get-Volume'   { Get-Volume })
    $Script:Volumes = @(foreach ($v in $w32vol) {
        $g = $gvol | Where-Object { $_.Path -eq $v.DeviceID } | Select-Object -First 1
        [pscustomobject]@{
            Nome = $v.Name; Rotulo = $v.Label; FS = $v.FileSystem; Cap = [double]$v.Capacity; Livre = [double]$v.FreeSpace
            Cluster = $v.BlockSize; Saude = $(if ($g) { [string]$g.HealthStatus } else { '' }); Letra = $v.DriveLetter
            PctLivre = (Pct $v.FreeSpace $v.Capacity)
        }
    })
    $pdisks = @(Coletar 'Armazenamento - Get-PhysicalDisk' { Get-PhysicalDisk })
    $disks  = @(Coletar 'Armazenamento - Get-Disk' { Get-Disk })
    $rel    = @(foreach ($p in $pdisks) { try { $r = $p | Get-StorageReliabilityCounter -ErrorAction Stop; if ($r) { [pscustomobject]@{ D = $p; R = $r } } } catch { } })

    # ============================================================================================
    #  MÁQUINAS VIRTUAIS
    # ============================================================================================
    Passo 'Máquinas virtuais'
    $vms = @(Coletar 'VMs - Get-VM' { Get-VM } | Where-Object { $_ } | Sort-Object Name)
    $replAll = @(Coletar 'Replicação - Get-VMReplication' { Get-VMReplication -ErrorAction SilentlyContinue })
    $usados = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $pastasVhd = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $Lista = New-Object System.Collections.Generic.List[object]
    $linhasVhd = New-Object System.Collections.Generic.List[object]

    foreach ($vm in $vms) {
        $n = $vm.Name
        $proc  = Coletar "VM $n - processador" { Get-VMProcessor -VM $vm }
        $mem   = Coletar "VM $n - memória" { Get-VMMemory -VM $vm }
        $hdds  = @(Coletar "VM $n - discos" { Get-VMHardDiskDrive -VM $vm })
        $snaps = @(Coletar "VM $n - checkpoints" { Get-VMSnapshot -VM $vm } | Where-Object { $_ })
        $nics  = @(Coletar "VM $n - adaptadores de rede" { Get-VMNetworkAdapter -VM $vm })
        $isv   = @(Coletar "VM $n - integration services" { Get-VMIntegrationService -VM $vm })
        $dvd   = @(Coletar "VM $n - DVD" { Get-VMDvdDrive -VM $vm })
        $fwv   = if ($vm.Generation -eq 2) { Coletar "VM $n - firmware" { Get-VMFirmware -VM $vm } } else { $null }
        $kvp   = if ($vm.State -eq 'Running') { Coletar "VM $n - KVP (sistema convidado)" { Get-GuestKvp $vm } } else { @{} }
        if (-not $kvp) { $kvp = @{} }

        $cadeiaBytes = 0.0; $discos = @()
        foreach ($d in $hdds) {
            $ctl = "{0} {1}:{2}" -f $d.ControllerType, $d.ControllerNumber, $d.ControllerLocation
            if ($null -ne $d.DiskNumber -and "$($d.DiskNumber)" -ne '') {
                $discos += [pscustomobject]@{ VM = $n; Controlador = $ctl; Caminho = "Disco físico $($d.DiskNumber) (pass-through)"; Tipo = 'Pass-through'; Formato = ''; Max = $null; Arquivo = $null; Frag = $null; Niveis = 0; Volume = $null; Ext = '' }
                continue
            }
            $p = $d.Path; $niv = 0; $topo = $null; $ultimo = $null
            while ($p -and $niv -lt 64) {
                [void]$usados.Add($p)
                $pai = Split-Path -Parent $p
                if ($pai) { [void]$pastasVhd.Add($pai) }
                $vh = Coletar "VM $n - Get-VHD $p" { Get-VHD -Path $p }
                if (-not $vh) { break }
                if (-not $topo) { $topo = $vh }
                $cadeiaBytes += [double]$vh.FileSize
                $niv++; $ultimo = $vh
                $p = $vh.ParentPath
            }
            if ($topo) {
                $discos += [pscustomobject]@{
                    VM = $n; Controlador = $ctl; Caminho = $topo.Path; Tipo = [string]$topo.VhdType; Formato = [string]$topo.VhdFormat
                    Max = [double]$topo.Size; Arquivo = [double]$topo.FileSize; Frag = $topo.FragmentationPercentage; Niveis = $niv
                    Volume = (VolumeDe $topo.Path); Ext = [IO.Path]::GetExtension($topo.Path).ToLower()
                    Base = $(if ($ultimo) { $ultimo.Path } else { '' }); BaseTipo = $(if ($ultimo) { [string]$ultimo.VhdType } else { '' })
                    Setor = ("{0}/{1}" -f $topo.LogicalSectorSize, $topo.PhysicalSectorSize)
                }
            } elseif ($d.Path) {
                $discos += [pscustomobject]@{ VM = $n; Controlador = $ctl; Caminho = $d.Path; Tipo = '(não lido)'; Formato = ''; Max = $null; Arquivo = $null; Frag = $null; Niveis = 0; Volume = (VolumeDe $d.Path); Ext = [IO.Path]::GetExtension($d.Path).ToLower() }
            }
        }
        foreach ($x in $discos) { $linhasVhd.Add($x) }
        $Lista.Add([pscustomobject]@{ VM = $vm; Proc = $proc; Mem = $mem; Discos = $discos; Snaps = $snaps; Nics = $nics; IS = $isv; Dvd = $dvd; Fw = $fwv; Kvp = $kvp; Bytes = $cadeiaBytes })
    }
    if ($vmh -and $vmh.VirtualHardDiskPath) { [void]$pastasVhd.Add($vmh.VirtualHardDiskPath.TrimEnd('\')) }

    # Arquivos de disco virtual nas pastas das VMs que não estão em uso por nenhuma VM.
    $orfaos = @(foreach ($pasta in $pastasVhd) {
        Get-ChildItem -LiteralPath $pasta -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.vhd', '.vhdx', '.avhd', '.avhdx' -and -not $usados.Contains($_.FullName) }
    })

    # ---- Verificações das VMs ------------------------------------------------------------------
    $Area = 'vms'; $C4 = '04 - Máquinas virtuais'
    $vcpuTot = 0; $vcpuOn = 0; $memStartTot = 0.0; $memAtrib = 0.0
    $vmOn = @($vms | Where-Object { $_.State -eq 'Running' })
    foreach ($i in $Lista) {
        $vm = $i.VM; $n = $vm.Name
        $vcpuTot += $vm.ProcessorCount
        if ($vm.State -eq 'Running') { $vcpuOn += $vm.ProcessorCount; $memAtrib += [double]$vm.MemoryAssigned }
        $memStartTot += [double]$vm.MemoryStartup

        $estado = [string]$vm.State
        $st = switch ($estado) { 'Running' { 'OK' } 'Off' { 'INFO' } default { 'WARN' } }
        Add-Check $Area $C4 "$n · Estado" $estado $st $(if ($estado -eq 'Off') { 'VM desligada. Confirmar se é esperado (VM inativa consome disco e licença).' } elseif ($st -eq 'WARN') { 'VM em estado salvo/pausado/transição; verificar.' } else { 'VM em execução.' })

        $opst = @($vm.OperationalStatus) | Select-Object -First 1
        if ($opst -and "$opst" -ne 'Ok') { Add-Check $Area $C4 "$n · Status operacional" ("$opst · " + $vm.Status) 'WARN' 'Hyper-V reporta condição anormal da VM.' }

        if ($vm.State -eq 'Running') {
            $hb = [string]$vm.Heartbeat
            $st = if ($hb -like 'Ok*') { 'OK' } elseif (-not $hb) { 'INFO' } else { 'WARN' }
            Add-Check $Area $C4 "$n · Heartbeat" $(if ($hb) { $hb } else { 'indisponível' }) $st 'Sem heartbeat o host não se comunica com o sistema convidado (integration services ausentes, desatualizados ou SO travado).'
            $memDem = [double]$vm.MemoryDemand; $memAs = [double]$vm.MemoryAssigned
            if ($vm.DynamicMemoryEnabled -and $memAs -gt 0) {
                $p = Pct $memDem $memAs
                $st = if ($p -ge 95) { 'WARN' } else { 'OK' }
                Add-Check $Area $C4 "$n · Pressão de memória" ("demanda {0} / atribuída {1} ({2}%)" -f (Bytes $memDem), (Bytes $memAs), (Num $p 0)) $st 'Memória dinâmica: demanda próxima do atribuído indica que a VM precisa de mais memória (ou do máximo configurado).'
            }
        }

        if ($verPadrao -and $vm.Version) {
            $st = 'OK'
            try { if ([version]$vm.Version -lt [version]$verPadrao) { $st = 'WARN' } } catch { }
            Add-Check $Area $C4 "$n · Versão de configuração" ("$($vm.Version) (host suporta $verPadrao)") $st 'Versão antiga não aproveita recursos do host. Atualização (Update-VMVersion) exige VM desligada, é irreversível e impede mover a VM para hosts mais antigos.'
        }

        $sa = [string]$vm.AutomaticStartAction
        $st = if ($sa -eq 'Nothing' -and $vm.State -eq 'Running') { 'WARN' } else { 'OK' }
        Add-Check $Area $C4 "$n · Ação ao iniciar o host" ("$sa · atraso $($vm.AutomaticStartDelay) s") $st 'Com "Nothing" a VM não volta sozinha após reinício do host. Use atraso para escalonar o boot (ex.: AD/DNS antes das aplicações).'
        $so = [string]$vm.AutomaticStopAction
        $st = switch ($so) { 'TurnOff' { 'WARN' } 'Save' { 'INFO' } default { 'OK' } }
        Add-Check $Area $C4 "$n · Ação ao desligar o host" $so $st $(if ($so -eq 'TurnOff') { 'TurnOff equivale a cortar a energia da VM: risco de corrupção. Usar ShutDown.' } elseif ($so -eq 'Save') { 'Save preserva o estado em disco (ocupa espaço do tamanho da RAM). ShutDown é preferível para servidores de aplicação/BD.' } else { 'Desligamento ordenado do sistema convidado.' })

        $ct = [string]$vm.CheckpointType
        $st = switch ($ct) { 'Standard' { 'WARN' } 'Disabled' { 'INFO' } default { 'OK' } }
        Add-Check $Area $C4 "$n · Tipo de checkpoint" $ct $st 'Production usa VSS no convidado (consistente para aplicações e bancos). Standard salva memória e estado e não é indicado para servidores de produção.'

        $iss = [string]$vm.IntegrationServicesState
        $st = if ($iss -match '(?i)required|necess|atualiz.*necess') { 'WARN' } else { 'INFO' }
        Add-Check $Area $C4 "$n · Integration Services" ("$($vm.IntegrationServicesVersion) $iss").Trim() $st 'Em convidados Windows Server 2016+ são atualizados pelo Windows Update; em Linux pelos pacotes do kernel (hv_utils/LIS).'
        $isOff = @($i.IS | Where-Object { -not $_.Enabled } | ForEach-Object { $_.Name })
        if ($isOff.Count) { Add-Check $Area $C4 "$n · Serviços de integração desabilitados" ($isOff -join ', ') 'INFO' 'Validar se a desativação é intencional (ex.: sincronização de hora desativada em controladores de domínio).' }

        foreach ($nic in $i.Nics) {
            if (-not $nic.SwitchName) { Add-Check $Area $C4 "$n · Adaptador '$($nic.Name)'" 'sem switch' 'WARN' 'Adaptador de rede não conectado a nenhum switch virtual.' }
        }
        $isos = @($i.Dvd | Where-Object { $_.Path })
        if ($isos.Count) { Add-Check $Area $C4 "$n · ISO montada" (($isos | ForEach-Object { $_.Path }) -join ', ') 'INFO' 'ISO montada impede Live Migration se o arquivo não estiver acessível no destino e pode afetar o boot.' }
        if ($i.Fw -and -not ($i.Fw.SecureBoot -eq 'On')) { Add-Check $Area $C4 "$n · Secure Boot" 'Desligado' 'INFO' 'VM geração 2 sem Secure Boot. Para Windows recomenda-se ligado (template MicrosoftWindows).' }
    }
    if ($vms.Count -eq 0) { Add-Check $Area $C4 'Máquinas virtuais' '0' 'INFO' 'Nenhuma VM encontrada neste host.' }
    if ($lp) {
        $ratio = [math]::Round($vcpuOn / $lp, 2)
        $st = if ($ratio -gt 4) { 'WARN' } else { 'OK' }
        Add-Check $Area $C4 'vCPU (VMs ligadas) : processadores lógicos' ("{0} : {1} ({2}:1)" -f $vcpuOn, $lp, (Num $ratio 2)) $st 'Proporção acima de 4:1 pode gerar espera por CPU (CPU Wait) em cargas intensas.'
        $maxv = ($vms | Measure-Object ProcessorCount -Maximum).Maximum
        if ($maxv -and $cores -and $sockets -and $maxv -gt ($cores / $sockets)) {
            Add-Check $Area $C4 'VM com mais vCPUs que cores de um socket' $maxv 'INFO' 'VMs maiores que um nó NUMA dependem de vNUMA; manter as VMs dentro de um socket quando possível.'
        }
    }
    if ($memTotal) {
        $reserva = 4GB
        $st = if ($memStartTot -gt ($memTotal - $reserva)) { 'WARN' } else { 'OK' }
        Add-Check $Area $C4 'Memória de inicialização de todas as VMs' ("{0} de {1} físicos" -f (Bytes $memStartTot), (Bytes $memTotal)) $st 'Soma da memória de inicialização de todas as VMs (ligadas e desligadas) comparada à memória física menos ~4 GB para o host. Acima disso nem todas as VMs conseguem ligar ao mesmo tempo.'
    }

    # ---- Verificações de armazenamento --------------------------------------------------------
    $Area = 'armazenamento'; $C5 = '05 - Armazenamento'
    $volVis = @($Script:Volumes | Where-Object { $_.Nome -notlike '\\?\*' -and $_.Cap -gt 1GB })
    foreach ($v in $volVis) {
        $st = if ($v.PctLivre -lt 10) { 'CRIT' } elseif ($v.PctLivre -lt 15) { 'WARN' } else { 'OK' }
        Add-Check $Area $C5 ("Espaço livre em {0}" -f $v.Nome) ("{0} de {1} ({2}%)" -f (Bytes $v.Livre), (Bytes $v.Cap), (Num $v.PctLivre 1)) $st 'Abaixo de 15% exige atenção; volume cheio pausa as VMs com discos dinâmicos ou checkpoints (estado "Paused-Critical").'
        if ($v.Saude -and $v.Saude -ne 'Healthy') { Add-Check $Area $C5 ("Saúde do volume {0}" -f $v.Nome) $v.Saude 'CRIT' 'Volume reporta problema de integridade.' }
    }
    foreach ($p in $pdisks) {
        $hs = [string]$p.HealthStatus; $os2 = (@($p.OperationalStatus) -join ', ')
        $st = if ($hs -eq 'Healthy' -and $os2 -match '^OK') { 'OK' } else { 'CRIT' }
        Add-Check $Area $C5 ("Disco físico {0} ({1})" -f $p.DeviceId, $p.FriendlyName) ("$hs · $os2") $st 'Saúde reportada pelo controlador/disco. Em RAID de hardware, validar também na ferramenta do fabricante (iDRAC, iLO, XClarity etc.).'
    }
    # Sobrecarga de discos dinâmicos/diferenciais por volume
    $porVol = $linhasVhd | Where-Object { $_.Volume -and $_.Tipo -in 'Dynamic', 'Differencing' } | Group-Object { $_.Volume.Nome }
    foreach ($g in $porVol) {
        $vol = $g.Group[0].Volume
        $cresc = ($g.Group | ForEach-Object { [math]::Max([double]0, $_.Max - $_.Arquivo) } | Measure-Object -Sum).Sum
        $st = if ($cresc -gt $vol.Livre) { 'WARN' } else { 'OK' }
        Add-Check $Area $C5 ("Crescimento potencial dos discos dinâmicos em {0}" -f $vol.Nome) ("{0} possíveis / {1} livres" -f (Bytes $cresc), (Bytes $vol.Livre)) $st 'Se os discos dinâmicos crescerem até o tamanho máximo, o volume pode esgotar e pausar as VMs. Monitorar ou converter para discos fixos.'
    }
    $emC = @($linhasVhd | Where-Object { $_.Caminho -and $_.Caminho.StartsWith($sysDrive, [StringComparison]::OrdinalIgnoreCase) })
    $st = if ($emC.Count) { 'WARN' } else { 'OK' }
    Add-Check $Area $C5 "Discos de VM no volume do sistema ($sysDrive)" $emC.Count $st 'Discos de VMs no volume do Windows competem por I/O e espaço com o host; se o C: encher, o host e as VMs param.'
    $dyn = @($linhasVhd | Where-Object { $_.Tipo -eq 'Dynamic' })
    Add-Check $Area $C5 'Discos dinâmicos' ("{0} de {1}" -f $dyn.Count, $linhasVhd.Count) 'INFO' 'Dinâmicos economizam espaço mas exigem monitoramento do volume. Fixos têm desempenho mais previsível.'
    $vhdOld = @($linhasVhd | Where-Object { $_.Formato -eq 'VHD' })
    if ($vhdOld.Count) { Add-Check $Area $C5 'Discos no formato VHD (legado)' $vhdOld.Count 'WARN' 'VHD é limitado a 2 TB e menos resiliente; converter para VHDX em janela de manutenção.' }
    $frag = @($linhasVhd | Where-Object { $_.Frag -ge 30 })
    if ($frag.Count) { Add-Check $Area $C5 'Discos virtuais com fragmentação >= 30%' $frag.Count 'INFO' 'Fragmentação alta em discos dinâmicos pode afetar desempenho em storage mecânico.' }
    $tamOrf = ($orfaos | Measure-Object Length -Sum).Sum
    $st = if ($orfaos.Count) { 'WARN' } else { 'OK' }
    Add-Check $Area $C5 'Arquivos de disco virtual não anexados' $(if ($orfaos.Count) { "{0} ({1})" -f $orfaos.Count, (Bytes $tamOrf) } else { '0' }) $st 'Arquivos .vhd/.vhdx/.avhdx nas pastas das VMs que não estão em uso. Podem ser sobras de VMs antigas ou checkpoints órfãos; validar antes de qualquer exclusão.'

    # ---- Checkpoints, replicação e backup ------------------------------------------------------
    $Area = 'protecao'; $C7 = '07 - Checkpoints'; $C8 = '08 - Replicação e backup'
    $todosSnaps = @($Lista | ForEach-Object { $_.Snaps })
    foreach ($i in $Lista) {
        $n = $i.VM.Name
        if ($i.Snaps.Count) {
            $velho = ($i.Snaps | Sort-Object CreationTime | Select-Object -First 1)
            $dias = [math]::Floor(($Script:Agora - $velho.CreationTime).TotalDays)
            $st = if ($dias -gt 30) { 'CRIT' } else { 'WARN' }
            Add-Check $Area $C7 "$n · Checkpoints" ("{0} (mais antigo: {1} dias)" -f $i.Snaps.Count, $dias) $st 'Checkpoint não é backup. A cadeia de .avhdx cresce continuamente, degrada o I/O e consome espaço. Remover após o uso (a remoção faz o merge automaticamente).'
            $rec = @($i.Snaps | Where-Object { "$($_.SnapshotType)" -match 'Recovery' })
            if ($rec.Count) { Add-Check $Area $C7 "$n · Checkpoints de recuperação (backup)" $rec.Count 'WARN' 'Checkpoints do tipo Recovery são criados pelo software de backup e deveriam ser removidos por ele ao fim do job. Remanescentes indicam falha no backup.' }
        } else {
            Add-Check $Area $C7 "$n · Checkpoints" '0' 'OK' 'Sem checkpoints.'
        }
        $avhd = @($i.Discos | Where-Object { $_.Ext -in '.avhdx', '.avhd' })
        if ($avhd.Count -and $i.Snaps.Count -eq 0) {
            Add-Check $Area $C7 "$n · Disco em uso é .avhdx sem checkpoint" $avhd.Count 'CRIT' 'A VM está rodando sobre um disco diferencial de checkpoint que não aparece no Hyper-V (merge pendente ou checkpoint órfão de backup). Requer correção planejada.'
        }
        $niv = ($i.Discos | Measure-Object Niveis -Maximum).Maximum
        if ($niv -gt 1) { Add-Check $Area $C7 "$n · Profundidade da cadeia de discos" $niv $(if ($niv -gt 3) { 'WARN' } else { 'INFO' }) 'Número de arquivos (base + diferenciais) que compõem o disco. Cada nível adiciona leitura extra.' }
    }

    $replSrv = Coletar 'Replicação - Get-VMReplicationServer' { Get-VMReplicationServer }
    if ($replAll.Count) {
        foreach ($r in $replAll) {
            $h = [string]$r.Health
            $st = switch ($h) { 'Normal' { 'OK' } 'Warning' { 'WARN' } 'Critical' { 'CRIT' } default { 'INFO' } }
            Add-Check $Area $C8 ("{0} · Hyper-V Replica" -f $r.VMName) ("{0} · {1} · {2}" -f $r.Mode, $r.State, $h) $st 'Saúde da replicação. Warning/Critical indicam atraso ou falha de replicação.'
        }
    } else {
        Add-Check $Area $C8 'Hyper-V Replica' 'não configurada' 'INFO' 'Nenhuma VM replicada. Em host standalone, a recuperação depende exclusivamente do backup.'
    }
    $vssTxt = Coletar 'Backup - vssadmin list writers' { (& vssadmin.exe list writers) -join "`n" }
    $writers = @()
    if ($vssTxt) {
        foreach ($bl in ($vssTxt -split "\n\s*\n")) {
            if ($bl -notmatch "'([^']+)'") { continue }
            $nomeW = $Matches[1]
            $linhas = @($bl -split "\n" | Where-Object { $_.Trim() })
            $estW = ($linhas | Where-Object { $_ -match '\[\d+\]' } | Select-Object -First 1)
            $num = if ($estW -match '\[(\d+)\]\s*(.*)$') { [int]$Matches[1] } else { 0.0 }
            $estTxt = if ($estW -match '\[(\d+)\]\s*(.*)$') { $Matches[2].Trim() } else { '' }
            $err = if ($linhas[-1] -match ':\s*(.+)$') { $Matches[1].Trim() } else { '' }
            $ok = ($num -eq 1 -and $err -match '(?i)no error|nenhum erro|sem erro')
            $writers += [pscustomobject]@{ 'Writer' = $nomeW; 'Estado' = "[$num] $estTxt"; 'Último erro' = $err; 'Status' = (Badge $(if ($ok) { 'OK' } else { 'WARN' }) $(if ($ok) { 'OK' } else { 'VERIFICAR' })); '_ok' = $ok }
        }
    }
    $wHv = $writers | Where-Object { $_.Writer -match 'Hyper-V' } | Select-Object -First 1
    if ($wHv) {
        Add-Check $Area $C8 'VSS Writer do Hyper-V' ("$($wHv.Estado) · $($wHv.'Último erro')") $(if ($wHv._ok) { 'OK' } else { 'WARN' }) 'O backup das VMs no nível do host depende do "Microsoft Hyper-V VSS Writer" estável e sem erro.'
    } elseif ($vssTxt) {
        Add-Check $Area $C8 'VSS Writer do Hyper-V' 'não listado' 'WARN' 'O writer do Hyper-V não apareceu em "vssadmin list writers".'
    }
    $wErr = @($writers | Where-Object { -not $_._ok })
    if ($wErr.Count) { Add-Check $Area $C8 'VSS writers com erro' (($wErr | ForEach-Object { $_.Writer }) -join ', ') 'WARN' 'Writers com falha podem fazer o backup falhar ou gerar cópias inconsistentes.' }
    $bkpRegex = 'Veeam|Acronis|Commvault|Arcserve|Backup Exec|Altaro|NAKIVO|Cohesity|Rubrik|Datto|Unitrends|Data Protection Manager|DPMRA|Bacula|Cobian|Iperius|Vembu|Zerto|Druva|NetBackup|Avamar|NetWorker|Spectrum Protect|Hornetsecurity|UrBackup|Synology Active Backup'
    $bkpSvc = @($svcAll | Where-Object { $_.DisplayName -match $bkpRegex -or $_.Name -match $bkpRegex })
    $wsb = $features | Where-Object { $_.Name -eq 'Windows-Server-Backup' -and $_.Installed }
    $wbs = if ($wsb -and (Get-Command Get-WBSummary -ErrorAction SilentlyContinue)) { Coletar 'Backup - Get-WBSummary' { Get-WBSummary } } else { $null }
    $txtBkp = @()
    if ($bkpSvc.Count) { $txtBkp += (($bkpSvc | ForEach-Object { $_.DisplayName } | Select-Object -Unique) -join ', ') }
    if ($wsb) { $txtBkp += 'Windows Server Backup' }
    Add-Check $Area $C8 'Software de backup identificado no host' $(if ($txtBkp.Count) { $txtBkp -join ' · ' } else { 'nenhum' }) $(if ($txtBkp.Count) { 'INFO' } else { 'WARN' }) 'Identificação por serviços instalados. Se o backup for feito dentro das VMs ou por servidor externo, validar com o cliente a política, a retenção e o último teste de restore.'
    if ($wbs) {
        $ult = $wbs.LastSuccessfulBackupTime
        $dias = if ($ult -and $ult.Year -gt 2000) { [math]::Floor(($Script:Agora - $ult).TotalDays) } else { $null }
        $st = if ($null -eq $dias) { 'WARN' } elseif ($dias -gt 2) { 'WARN' } else { 'OK' }
        Add-Check $Area $C8 'Windows Server Backup · último sucesso' $(if ($null -ne $dias) { "$(Dt $ult) (há $dias dias)" } else { 'nunca' }) $st 'Último backup bem-sucedido do Windows Server Backup.'
    }

    # ============================================================================================
    #  REDE
    # ============================================================================================
    Passo 'Rede'
    $Area = 'rede'; $C6 = '06 - Rede'
    $switches = @(Coletar 'Rede - Get-VMSwitch' { Get-VMSwitch })
    $nicsAll  = @(Coletar 'Rede - Get-NetAdapter' { Get-NetAdapter })
    $nicsPhy  = @($nicsAll | Where-Object { $_.HardwareInterface -or $_.InterfaceDescription -match 'Multiplexor' })
    $vmq      = @(Coletar 'Rede - Get-NetAdapterVmq' { Get-NetAdapterVmq -ErrorAction SilentlyContinue })
    $teams    = @(Coletar 'Rede - Get-NetLbfoTeam' { Get-NetLbfoTeam -ErrorAction SilentlyContinue })
    $ipcfg    = @(Coletar 'Rede - Get-NetIPConfiguration' { Get-NetIPConfiguration })
    $hostVnic = @(Coletar 'Rede - adaptadores virtuais do host' { Get-VMNetworkAdapter -ManagementOS })
    $uplinkDesc = @()
    foreach ($s in $switches) {
        $descs = @(PV $s 'NetAdapterInterfaceDescriptions', 'NetAdapterInterfaceDescription') | Where-Object { $_ }
        $uplinkDesc += $descs
        if ("$($s.SwitchType)" -eq 'External') {
            foreach ($d in $descs) {
                $a = $nicsAll | Where-Object { $_.InterfaceDescription -eq $d } | Select-Object -First 1
                if ($a) {
                    $st = if ($a.Status -ne 'Up') { 'CRIT' } elseif ([double]$a.Speed -lt 1e9) { 'WARN' } else { 'OK' }
                    Add-Check $Area $C6 ("Uplink do switch '{0}'" -f $s.Name) ("{0} · {1} · {2}" -f $a.Name, $a.Status, $a.LinkSpeed) $st 'Adaptador físico do switch externo deve estar Up e com 1 Gbps ou mais.'
                }
            }
            $vms2 = @($Lista | Where-Object { $_.Nics | Where-Object { $_.SwitchName -eq $s.Name } }).Count
            Add-Check $Area $C6 ("Switch '{0}'" -f $s.Name) ("Externo · gerenciamento: {0} · SET: {1} · {2} VM(s)" -f (SN $s.AllowManagementOS), (SN $s.EmbeddedTeamingEnabled), $vms2) 'INFO' 'Switch externo ligado à rede física.'
        } else {
            Add-Check $Area $C6 ("Switch '{0}'" -f $s.Name) ("{0}" -f $s.SwitchType) 'INFO' 'Switch Interno/Privado: as VMs conectadas não acessam a rede física por ele.'
        }
    }
    if ($switches.Count -eq 0) { Add-Check $Area $C6 'Switches virtuais' '0' 'WARN' 'Nenhum switch virtual configurado.' }
    if ($teams.Count) { Add-Check $Area $C6 'NIC Teaming (LBFO)' (($teams | ForEach-Object { "$($_.Name) ($($_.TeamingMode)/$($_.LoadBalancingAlgorithm), $($_.Status))" }) -join '; ') $(if ($teams | Where-Object { "$($_.Status)" -ne 'Up' }) { 'WARN' } else { 'INFO' }) 'A partir do Windows Server 2016 o recomendado para Hyper-V é Switch Embedded Teaming (SET); LBFO com switch virtual foi descontinuado no Windows Server 2022.' }
    foreach ($q in $vmq) {
        $a = $nicsAll | Where-Object { $_.Name -eq $q.Name } | Select-Object -First 1
        if ($q.Enabled -and $a -and $a.Status -eq 'Up' -and [double]$a.Speed -le 1e9 -and $a.InterfaceDescription -match 'Broadcom') {
            Add-Check $Area $C6 ("VMQ em '{0}'" -f $q.Name) ("Habilitado · {0} · {1}" -f $a.LinkSpeed, $a.InterfaceDescription) 'WARN' 'Problema conhecido: VMQ habilitado em placas Broadcom de 1 Gbps causa perda de rede/lentidão nas VMs. Recomenda-se desabilitar VMQ nessas placas ou atualizar o driver.'
        }
    }
    foreach ($a in @($nicsAll | Where-Object { $_.HardwareInterface })) {
        $dd = $null; try { if ($a.DriverDate) { $dd = [datetime]$a.DriverDate } } catch { }
        if ($dd -and ($Script:Agora - $dd).TotalDays -gt 3 * 365 -and $a.Status -eq 'Up') {
            Add-Check $Area $C6 ("Driver da placa '{0}'" -f $a.Name) ("{0} · {1}" -f $a.DriverVersion, $dd.ToString('dd/MM/yyyy')) 'WARN' 'Driver com mais de 3 anos. Atualizar driver e firmware pelo pacote do fabricante do servidor.'
        }
    }
    $nicsDown = @($nicsAll | Where-Object { $_.HardwareInterface -and $_.Status -ne 'Up' })
    if ($nicsDown.Count) { Add-Check $Area $C6 'Adaptadores físicos desconectados' (($nicsDown | ForEach-Object { $_.Name }) -join ', ') 'INFO' 'Placas sem link. Confirmar se é esperado (portas não cabeadas).' }

    # ============================================================================================
    #  EVENTOS
    # ============================================================================================
    Passo "Eventos (últimos $DiasEventos dias)"
    $Area = 'eventos'; $C9 = '09 - Eventos'
    $logsHv = @(Coletar 'Eventos - lista de logs Hyper-V' { Get-WinEvent -ListLog 'Microsoft-Windows-Hyper-V*' -ErrorAction SilentlyContinue } |
                Where-Object { $_ -and $_.LogName -like '*-Admin' -and $_.RecordCount -gt 0 })
    $evHv = @(foreach ($l in $logsHv) { Get-Eventos $l.LogName @(1, 2, 3) $desde })
    $evSys = @(Get-Eventos 'System' @(1, 2) $desde)
    $evApp = @(Get-Eventos 'Application' @(1, 2) $desde)
    $evDesl = @($evSys | Where-Object {
        ($_.ProviderName -eq 'Microsoft-Windows-Kernel-Power' -and $_.Id -eq 41) -or ($_.ProviderName -eq 'EventLog' -and $_.Id -eq 6008) })
    $evDisco = @($evSys + (Get-Eventos 'System' @(3) $desde) | Where-Object {
        ($_.ProviderName -match '^(disk|Disk)$' -and $_.Id -in 7, 11, 51, 52, 153) -or
        ($_.ProviderName -match 'Ntfs' -and $_.Id -in 55, 98, 137, 140) -or
        ($_.ProviderName -match 'storport|stornvme|iaStor|megasas|percsas|HpCISSs|smartpqi' -and $_.Id -in 129, 9, 11) })

    $hvErr = @($evHv | Where-Object { $_.Level -in 1, 2 })
    $st = if ($hvErr | Where-Object { $_.Level -eq 1 }) { 'CRIT' } elseif ($hvErr.Count) { 'WARN' } else { 'OK' }
    Add-Check $Area $C9 'Erros/críticos nos logs do Hyper-V' $hvErr.Count $st "Logs Microsoft-Windows-Hyper-V-*-Admin nos últimos $DiasEventos dias. Ver detalhamento por ID."
    $hvWarn = @($evHv | Where-Object { $_.Level -eq 3 })
    Add-Check $Area $C9 'Avisos nos logs do Hyper-V' $hvWarn.Count 'INFO' 'Avisos costumam indicar problemas de integration services, checkpoints ou rede das VMs.'
    $st = if ($evDisco.Count) { 'CRIT' } else { 'OK' }
    Add-Check $Area $C9 'Erros de disco / sistema de arquivos / controladora' $evDisco.Count $st 'Eventos disk 7/11/51/153, Ntfs 55/98/137/140 e storport 129 indicam falha de mídia, timeout ou corrupção. Verificar storage e controladora imediatamente.'
    $st = if ($evDesl.Count) { 'WARN' } else { 'OK' }
    Add-Check $Area $C9 'Desligamentos inesperados do host' $evDesl.Count $st 'Kernel-Power 41 / EventLog 6008: o host reiniciou sem desligamento ordenado (queda de energia, travamento ou tela azul).'
    Add-Check $Area $C9 'Erros/críticos no log System' $evSys.Count $(if ($evSys | Where-Object { $_.Level -eq 1 }) { 'WARN' } else { 'INFO' }) 'Total no período; ver os IDs mais frequentes no detalhamento.'
    Add-Check $Area $C9 'Erros/críticos no log Application' $evApp.Count 'INFO' 'Total no período; ver os IDs mais frequentes no detalhamento.'

    # ============================================================================================
    #  PERFORMANCE (amostra)
    # ============================================================================================
    Passo "Amostra de desempenho ($AmostraSegundos s)"
    $Area = 'performance'; $C10 = '10 - Desempenho'
    $qLP = { Get-CimInstance Win32_PerfRawData_HvStats_HyperVHypervisorLogicalProcessor }
    $qRoot = { Get-CimInstance Win32_PerfRawData_HvStats_HyperVHypervisorRootVirtualProcessor }
    $qDisk = { Get-CimInstance Win32_PerfRawData_PerfDisk_LogicalDisk }
    $lp0 = @(Coletar 'Desempenho - contadores do hipervisor' $qLP); $rt0 = @(Coletar 'Desempenho - root VP' $qRoot); $dk0 = @(Coletar 'Desempenho - discos' $qDisk)
    Start-Sleep -Seconds ([math]::Max(3, $AmostraSegundos))
    $lp1 = @(Coletar 'Desempenho - contadores do hipervisor' $qLP); $rt1 = @(Coletar 'Desempenho - root VP' $qRoot); $dk1 = @(Coletar 'Desempenho - discos' $qDisk)
    $mediaTimer = {
        param($a0, $a1)
        $vals = foreach ($x1 in ($a1 | Where-Object { $_.Name -ne '_Total' })) {
            $x0 = $a0 | Where-Object { $_.Name -eq $x1.Name } | Select-Object -First 1
            if ($x0) {
                $dt = [double]$x1.Timestamp_Sys100NS - [double]$x0.Timestamp_Sys100NS
                if ($dt -gt 0) { [math]::Max([double]0, [math]::Min([double]100, 100.0 * ([double]$x1.PercentTotalRunTime - [double]$x0.PercentTotalRunTime) / $dt)) }
            }
        }
        $vals = @($vals)
        if ($vals.Count) { [pscustomobject]@{ Media = ($vals | Measure-Object -Average).Average; Max = ($vals | Measure-Object -Maximum).Maximum; Qtd = $vals.Count } }
    }
    $cpuLP = if ($lp0.Count -and $lp1.Count) { & $mediaTimer $lp0 $lp1 } else { $null }
    $cpuRoot = if ($rt0.Count -and $rt1.Count) { & $mediaTimer $rt0 $rt1 } else { $null }
    if ($cpuLP) {
        $st = if ($cpuLP.Media -gt 85) { 'CRIT' } elseif ($cpuLP.Media -gt 70) { 'WARN' } else { 'OK' }
        Add-Check $Area $C10 'CPU total do hipervisor (média da amostra)' ("{0}% (LP mais ocupado: {1}%)" -f (Num $cpuLP.Media 1), (Num $cpuLP.Max 1)) $st "Hyper-V Hypervisor Logical Processor\% Total Run Time em $AmostraSegundos s. É o uso real da CPU física (o Gerenciador de Tarefas do host não mostra o consumo das VMs)."
    } else { Nota 'Desempenho' 'Contadores do hipervisor (HvStats) indisponíveis.' }
    if ($cpuRoot) {
        $st = if ($cpuRoot.Media -gt 50) { 'WARN' } else { 'OK' }
        Add-Check $Area $C10 'CPU da partição raiz (host)' ("{0}%" -f (Num $cpuRoot.Media 1)) $st 'Consumo do próprio Windows do host. Deve ser baixo; o host não deve rodar outras funções (SQL, AD, arquivos etc.).'
    }
    $latencias = @()
    foreach ($d1 in ($dk1 | Where-Object { $_.Name -ne '_Total' -and $_.Name -notmatch '^HarddiskVolume' })) {
        $d0 = $dk0 | Where-Object { $_.Name -eq $d1.Name } | Select-Object -First 1
        if (-not $d0) { continue }
        $f = [double]$d1.Frequency_PerfTime; $dt = ([double]$d1.Timestamp_PerfTime - [double]$d0.Timestamp_PerfTime) / $f
        $br = [double]$d1.AvgDisksecPerRead_Base - [double]$d0.AvgDisksecPerRead_Base
        $bw = [double]$d1.AvgDisksecPerWrite_Base - [double]$d0.AvgDisksecPerWrite_Base
        $lr = if ($br -gt 0) { 1000.0 * (([double]$d1.AvgDisksecPerRead - [double]$d0.AvgDisksecPerRead) / $f) / $br } else { 0.0 }
        $lw = if ($bw -gt 0) { 1000.0 * (([double]$d1.AvgDisksecPerWrite - [double]$d0.AvgDisksecPerWrite) / $f) / $bw } else { 0.0 }
        $latencias += [pscustomobject]@{
            'Volume' = $d1.Name; 'Leituras/s' = (Num ($br / $dt) 1); 'Escritas/s' = (Num ($bw / $dt) 1)
            'Latência leitura (ms)' = (Num $lr 1); 'Latência escrita (ms)' = (Num $lw 1); 'Fila atual' = $d1.CurrentDiskQueueLength
            '_lr' = $lr; '_lw' = $lw
        }
    }
    foreach ($l in $latencias) {
        $m = [math]::Max($l._lr, $l._lw)
        $st = if ($m -gt 50) { 'CRIT' } elseif ($m -gt 20) { 'WARN' } else { 'OK' }
        Add-Check $Area $C10 ("Latência de disco em {0}" -f $l.Volume) ("leitura {0} ms · escrita {1} ms" -f $l.'Latência leitura (ms)', $l.'Latência escrita (ms)') $st 'Amostra curta e indicativa. Referência: até 10 ms bom, 10-20 ms aceitável, acima de 20 ms lento. Confirmar com monitoramento contínuo.'
    }
    foreach ($vm in $vmOn) {
        Add-Check $Area $C10 ("{0} · CPU no momento" -f $vm.Name) ("{0}% de {1} vCPU" -f $vm.CPUUsage, $vm.ProcessorCount) 'INFO' 'Get-VM CPUUsage: uso instantâneo da VM em relação aos vCPUs atribuídos.'
    }

    # ============================================================================================
    #  HTML
    # ============================================================================================
    Passo 'Gerando HTML'
    $abas = [ordered]@{ geral = 'Visão Geral'; atencao = 'Atenção'; vms = 'Máquinas Virtuais'; armazenamento = 'Armazenamento'; rede = 'Rede'; protecao = 'Checkpoints e Backup'; eventos = 'Eventos'; performance = 'Desempenho'; checklist = 'Checklist' }
    $nCrit = @($Script:Checks | Where-Object { $_.Status -eq 'CRIT' }).Count
    $nWarn = @($Script:Checks | Where-Object { $_.Status -eq 'WARN' }).Count
    $nOk   = @($Script:Checks | Where-Object { $_.Status -eq 'OK' }).Count
    $nInfo = @($Script:Checks | Where-Object { $_.Status -eq 'INFO' }).Count
    $conf = if (($nOk + $nWarn + $nCrit) -gt 0) { [math]::Round(100.0 * $nOk / ($nOk + $nWarn + $nCrit)) } else { 100 }

    $out = New-Object System.Text.StringBuilder
    $titulo = "Assessment Hyper-V · $hostNome" + $(if ($Cliente) { " · $Cliente" } else { '' })
    [void]$out.Append("<!DOCTYPE html>`r`n" + $Script:Cabecalho.Replace('{{TITULO}}', (Esc $titulo)))
    [void]$out.Append('<style>details>p.empty{margin:0 16px}.vmcards{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.vmcards .card h3{display:flex;justify-content:space-between;align-items:center}@media (max-width:900px){.vmcards{grid-template-columns:1fr}}</style>')

    # --- Capa
    $hvTxt = if ($verPadrao) { "Hyper-V · config. VM $verPadrao" } else { 'Hyper-V' }
    $meta = [ordered]@{
        'Data da coleta' = (Dt $Script:Agora)
        'Sistema operacional' = ("$osNome").Replace('Microsoft ', '')
        'Build' = ("$build" + $(if ($release) { " ($release)" } else { '' }))
        'Virtualização' = $hvTxt
        'Fabricante / modelo' = ("$(PV $cs 'Manufacturer') $(PV $cs 'Model')").Trim()
        'Domínio' = $(if ($cs -and $cs.PartOfDomain) { $cs.Domain } else { 'Workgroup' })
        'Hardware' = ("{0} LP · {1}" -f $lp, (Bytes $memTotal 0))
        'Script de coleta' = "$($Script:Versao) (somente leitura)"
    }
    [void]$out.Append('<header class="cover"><div class="kicker">Assessment técnico · Hyper-V</div><h1>Levantamento do Ambiente de Virtualização</h1>')
    [void]$out.Append('<div class="client">Cliente: <b>' + (Esc $(if ($Cliente) { $Cliente } else { '-' })) + '</b></div>')
    [void]$out.Append('<div class="server"><span class="lbl">Host Hyper-V</span><span class="val">' + (Esc $fqdn) + '</span></div><div class="meta">')
    foreach ($k in $meta.Keys) { [void]$out.Append('<div><span>' + (Esc $k) + '</span><b>' + (Esc $meta[$k]) + '</b></div>') }
    [void]$out.Append('</div></header>')

    # --- Guias
    [void]$out.Append('<nav class="tabs">')
    foreach ($k in $abas.Keys) {
        $c = @($Script:Checks | Where-Object { ($_.Area -eq $k -or $k -eq 'atencao') -and $_.Status -in 'CRIT', 'WARN' })
        $badge = ''
        if ($c.Count -and $k -ne 'checklist') {
            $cl = if ($c | Where-Object { $_.Status -eq 'CRIT' }) { 'crit' } else { 'warn' }
            $badge = '<span class="n ' + $cl + '">' + $c.Count + '</span>'
        }
        [void]$out.Append('<button data-t="' + $k + '" onclick="tab(''' + $k + ''')" type="button">' + (Esc $abas[$k]) + $badge + '</button>')
    }
    [void]$out.Append('</nav>')

    # --- Visão Geral
    $cor = if ($conf -ge 80) { '#2E7D4F' } elseif ($conf -ge 60) { '#A86400' } else { '#B3121B' }
    $arco = [math]::Round(326.73 * $conf / 100, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
    $snTot = $todosSnaps.Count
    $snVelho = if ($snTot) { [math]::Floor(($Script:Agora - ($todosSnaps | Sort-Object CreationTime | Select-Object -First 1).CreationTime).TotalDays) } else { 0.0 }
    $vmBytes = ($Lista | Measure-Object Bytes -Sum).Sum
    $volVm = @($linhasVhd | Where-Object { $_.Volume } | ForEach-Object { $_.Volume.Nome } | Select-Object -Unique)
    [void]$out.Append('<div class="tabpane" id="t-geral"><section id="resumo"><h2><span class="num">01</span>Resumo executivo</h2><div class="exec"><div class="gauge"><svg viewbox="0 0 140 140"><circle cx="70" cy="70" fill="none" r="52" stroke="#E4E4E1" stroke-width="14"></circle><circle cx="70" cy="70" fill="none" r="52" stroke="' + $cor + '" stroke-dasharray="' + $arco + ' 326.73" stroke-linecap="round" stroke-width="14" transform="rotate(-90 70 70)"></circle><text class="pct" text-anchor="middle" x="70" y="80">' + $conf + '%</text></svg><div class="cap">das verificações automáticas em conformidade</div></div>')
    [void]$out.Append('<div class="counts"><div class="cnt crit"><span>Crítico</span><b>' + $nCrit + '</b><em>ação imediata</em></div><div class="cnt warn"><span>Atenção</span><b>' + $nWarn + '</b><em>revisar</em></div><div class="cnt ok"><span>OK</span><b>' + $nOk + '</b><em>em conformidade</em></div><div class="cnt"><span>Info</span><b>' + $nInfo + '</b><em>informativo</em></div></div></div><div class="kpis">')
    [void]$out.Append((Kpi 'Processadores lógicos' $lp ("{0} socket(s) · {1} cores" -f $sockets, $cores)))
    [void]$out.Append((Kpi 'Memória física' (Bytes $memTotal 1) ("livre: {0} ({1}%)" -f (Bytes $memLivre), (Num $pctLivre 0))))
    [void]$out.Append((Kpi 'Máquinas virtuais' $vms.Count ("{0} ligada(s) · {1} desligada(s)" -f $vmOn.Count, ($vms.Count - $vmOn.Count))))
    [void]$out.Append((Kpi 'vCPUs alocadas' $vcpuTot ("{0} em VMs ligadas · {1}:1 sobre LPs" -f $vcpuOn, $(if ($lp) { Num ($vcpuOn / $lp) 2 } else { '-' }))))
    [void]$out.Append((Kpi 'Memória das VMs' (Bytes $memAtrib 1) ("atribuída às VMs ligadas · {0}% do host" -f (Num (Pct $memAtrib $memTotal) 0))))
    [void]$out.Append((Kpi 'Discos das VMs' (Bytes $vmBytes 1) ("{0} disco(s) · em {1}" -f $linhasVhd.Count, ($volVm -join ', '))))
    [void]$out.Append((Kpi 'Checkpoints' $snTot $(if ($snTot) { "mais antigo com $snVelho dias" } else { 'nenhum checkpoint' })))
    [void]$out.Append((Kpi 'Uptime do host' $(if ($uptime) { "$([math]::Floor($uptime.TotalDays)) dias" } else { '-' }) $(if ($boot) { "desde $(Dt $boot)" } else { '' })))
    [void]$out.Append('</div></section><section>' + (ChecksArea 'geral') + (Ferramentas))

    $kvHost = [ordered]@{
        'Nome' = $hostNome; 'FQDN' = $fqdn; 'Domínio / grupo' = $(if ($cs) { $cs.Domain } else { '' }); 'Fabricante' = (PV $cs 'Manufacturer'); 'Modelo' = (PV $cs 'Model')
        'Número de série' = (PV $bios 'SerialNumber'); 'BIOS' = ("{0} ({1})" -f (PV $bios 'SMBIOSBIOSVersion'), (Dt (PV $bios 'ReleaseDate')))
        'Processador' = $(if ($cpus) { $cpus[0].Name.Trim() } else { '' }); 'Sockets / cores / lógicos' = "$sockets / $cores / $lp"
        'Virtualização no firmware (SLAT)' = $(if ($cpus) { SN $cpus[0].SecondLevelAddressTranslationExtensions } else { '' })
        'Memória física' = (Bytes $memTotal); 'Memória livre' = (Bytes $memLivre); 'Módulos de memória' = $dimm.Count
        'Arquivo de paginação' = (($pf | ForEach-Object { "$($_.Name) ($(Num $_.AllocatedBaseSize) MB)" }) -join ', ')
    }
    $kvOs = [ordered]@{
        'Sistema operacional' = $osNome; 'Edição' = (PV $nt 'EditionID'); 'Versão' = (PV $os 'Version'); 'Build (UBR)' = $build; 'Release' = $release
        'Instalado em' = (Dt (PV $os 'InstallDate')); 'Último boot' = (Dt $boot); 'Uptime' = (Span $uptime)
        'Idioma (LCID)' = (PV $os 'OSLanguage'); 'Plano de energia' = $plano; 'Fonte de hora' = $fonteHora
        'Firewall' = (($fw | ForEach-Object { "$($_.Name): $(SN $_.Enabled)" }) -join ' · ')
        'Reinício pendente' = $(if ($motivos.Count) { $motivos -join ', ' } else { 'NÃO' })
    }
    [void]$out.Append((Det '01' 'Host físico' 'Hardware, BIOS e memória do servidor.' (KV $kvHost) $null))
    [void]$out.Append((Det '02' 'Sistema operacional' 'Versão, build, boot, energia, hora e firewall.' (KV $kvOs) $null))
    $memRows = foreach ($m in $dimm) { [pscustomobject]@{ 'Slot' = $m.DeviceLocator; 'Banco' = $m.BankLabel; 'Capacidade' = (Bytes $m.Capacity 0); 'Velocidade (MHz)' = $m.Speed; 'Fabricante' = $m.Manufacturer; 'Part number' = ("$($m.PartNumber)").Trim() } }
    [void]$out.Append((Det '01' 'Módulos de memória' 'Win32_PhysicalMemory.' (Tbl $memRows 'Slot', 'Banco', 'Capacidade', 'Velocidade (MHz)', 'Fabricante', 'Part number' -Num 'Capacidade', 'Velocidade (MHz)' -Data) $dimm.Count))
    if ($vmh) {
        $kvHv = [ordered]@{
            'Processadores lógicos' = $vmh.LogicalProcessorCount; 'Memória (capacidade)' = (Bytes $vmh.MemoryCapacity)
            'Versão padrão de configuração' = $verPadrao; 'Pasta padrão de VHDs' = $vmh.VirtualHardDiskPath; 'Pasta padrão de VMs' = $vmh.VirtualMachinePath
            'NUMA spanning' = (SN $vmh.NumaSpanningEnabled); 'Enhanced Session Mode' = (SN $vmh.EnableEnhancedSessionMode)
            'Live Migration' = (SN $vmh.VirtualMachineMigrationEnabled); 'Migrações simultâneas' = $vmh.MaximumVirtualMachineMigrations
            'Autenticação de migração' = $vmh.VirtualMachineMigrationAuthenticationType; 'Faixa de MAC' = ("{0} - {1}" -f $vmh.MacAddressMinimum, $vmh.MacAddressMaximum)
            'Hyper-V Replica (servidor de réplica)' = $(if ($replSrv) { SN $replSrv.ReplicationEnabled } else { '' })
        }
        [void]$out.Append((Det '03' 'Configuração do host Hyper-V' 'Get-VMHost.' (KV $kvHv) $null))
    }
    $featRows = foreach ($f in $features) { [pscustomobject]@{ 'Recurso' = $f.Name; 'Descrição' = $f.DisplayName; 'Estado' = (Badge $(if ($f.Installed) { 'OK' } else { 'INFO' }) ([string]$f.InstallState)) } }
    [void]$out.Append((Det '03' 'Funções e recursos relevantes' 'Get-WindowsFeature.' (Tbl $featRows 'Recurso', 'Descrição', 'Estado' -Raw 'Estado' -Data) @($features).Count))
    $hvSvc = foreach ($s in ($svcAll | Where-Object { $_.Name -match '^(vmms|vmcompute|hvhost|nvspwmi|vhdsvc|W32Time|WinRM|VSS|ClusSvc)$' })) {
        [pscustomobject]@{ 'Serviço' = $s.Name; 'Descrição' = $s.DisplayName; 'Início' = $s.StartMode; 'Estado' = (Badge $(if ($s.State -eq 'Running') { 'OK' } elseif ($s.StartMode -eq 'Auto') { 'WARN' } else { 'INFO' }) $s.State); 'Conta' = $s.StartName }
    }
    [void]$out.Append((Det '03' 'Serviços do Hyper-V e de suporte' 'Estado dos serviços usados pelo Hyper-V, hora, gerenciamento e VSS.' (Tbl $hvSvc 'Serviço', 'Descrição', 'Início', 'Estado', 'Conta' -Raw 'Estado' -Data) @($hvSvc).Count))
    $ignorar = '^(gupdate|gupdatem|sppsvc|RemoteRegistry|MapsBroker|edgeupdate|edgeupdatem|ShellHWDetection|tiledatamodelsvc|WbioSrvc|CDPSvc|TrustedInstaller|BITS|wuauserv|sedsvc|DoSvc|UsoSvc|WaaSMedicSvc|clr_optimization.*|OneSyncSvc_.*|CDPUserSvc_.*|WpnUserService_.*|GoogleUpdater.*|AdobeARMservice|wscsvc|stisvc|dmwappushservice)$'
    $svcPar = @($svcAll | Where-Object { $_.StartMode -eq 'Auto' -and $_.State -ne 'Running' -and $_.Name -notmatch $ignorar })
    $rows = foreach ($s in $svcPar) { [pscustomobject]@{ 'Serviço' = $s.Name; 'Descrição' = $s.DisplayName; 'Estado' = $s.State; 'Código de saída' = $s.ExitCode } }
    [void]$out.Append((Det '02' 'Serviços automáticos parados' 'Serviços com início automático que não estão em execução (exceto os que param sozinhos por design).' (Tbl $rows 'Serviço', 'Descrição', 'Estado', 'Código de saída' -Num 'Código de saída' -Data) $svcPar.Count))
    $rows = foreach ($h in ($hf | Sort-Object { if ($_.InstalledOn) { [datetime]$_.InstalledOn } else { [datetime]::MinValue } } -Descending | Select-Object -First 20)) { [pscustomobject]@{ 'KB' = $h.HotFixID; 'Tipo' = $h.Description; 'Instalado em' = $(if ($h.InstalledOn) { ([datetime]$h.InstalledOn).ToString('dd/MM/yyyy') } else { '' }); 'Instalado por' = $h.InstalledBy } }
    [void]$out.Append((Det '02' 'Atualizações instaladas (20 mais recentes)' 'Get-HotFix.' (Tbl $rows 'KB', 'Tipo', 'Instalado em', 'Instalado por' -Data) "$($hf.Count) no total"))
    if ($mp) {
        $kvAv = [ordered]@{
            'Serviço antimalware' = (SN $mp.AMServiceEnabled); 'Proteção em tempo real' = (SN $mp.RealTimeProtectionEnabled); 'Versão do produto' = $mp.AMProductVersion
            'Assinaturas atualizadas em' = (Dt $mp.AntivirusSignatureLastUpdated); 'Exclusões automáticas desativadas' = $(if ($mpPref) { SN $mpPref.DisableAutoExclusions } else { '' })
            'Exclusões de pasta' = $(if ($mpPref) { (@($mpPref.ExclusionPath) -join '; ') } else { '' }); 'Exclusões de processo' = $(if ($mpPref) { (@($mpPref.ExclusionProcess) -join '; ') } else { '' })
            'Exclusões de extensão' = $(if ($mpPref) { (@($mpPref.ExclusionExtension) -join '; ') } else { '' })
        }
        [void]$out.Append((Det '02' 'Microsoft Defender' 'Get-MpComputerStatus / Get-MpPreference.' (KV $kvAv) $null))
    }
    [void]$out.Append('</section></div>')

    # --- Atenção
    $ord = @{ CRIT = 0; WARN = 1 }
    $rows = foreach ($x in ($Script:Checks | Where-Object { $_.Status -in 'CRIT', 'WARN' } | Sort-Object @{ Expression = { $ord[$_.Status] } }, Cat)) {
        [pscustomobject]@{ 'Status' = (Badge $x.Status); 'Categoria' = $x.Cat; 'Indicador' = $x.Ind; 'Valor' = $x.Val; 'Observação' = $x.Obs }
    }
    [void]$out.Append('<div class="tabpane" id="t-atencao"><section id="atencao"><h2><span class="num">02</span>Pontos de atenção</h2><p class="lead">Itens classificados automaticamente como CRÍTICO ou ATENÇÃO. A classificação é indicativa e deve ser validada com o contexto do ambiente e do negócio.</p>')
    [void]$out.Append($(if (@($rows).Count) { Tbl $rows 'Status', 'Categoria', 'Indicador', 'Valor', 'Observação' -Raw 'Status' -Num 'Valor' } else { '<p class="empty">Nenhum ponto de atenção identificado.</p>' }))
    [void]$out.Append('</section></div>')

    # --- Máquinas virtuais
    [void]$out.Append('<div class="tabpane" id="t-vms"><section><h2><span class="num">03</span>Máquinas virtuais</h2><p class="lead">Inventário, recursos, integração com o host e configuração de cada VM.</p>' + (ChecksArea 'vms'))
    [void]$out.Append('<div class="sub-h">Recursos alocados</div><div class="grid2">')
    $bm = foreach ($vm in $vms) {
        $v = if ($vm.State -eq 'Running') { [double]$vm.MemoryAssigned } else { [double]$vm.MemoryStartup }
        [pscustomobject]@{ N = $vm.Name; P = (Pct $v $memTotal); C = 'db'; T = ("{0} · {1}% do host" -f (Bytes $v), (Num (Pct $v $memTotal) 0)) + $(if ($vm.State -ne 'Running') { ' (desligada)' } else { '' }) }
    }
    [void]$out.Append((Card 'Memória por VM (em relação ao host)' (Bars $bm)))
    $bc = foreach ($vm in $vms) { [pscustomobject]@{ N = $vm.Name; P = (Pct $vm.ProcessorCount $lp); C = 'db'; T = ("{0} vCPU · {1}% em uso" -f $vm.ProcessorCount, $vm.CPUUsage) } }
    [void]$out.Append((Card 'vCPUs por VM (em relação aos processadores lógicos)' (Bars $bc)))
    [void]$out.Append('</div>' + (Ferramentas))
    $rows = foreach ($i in $Lista) {
        $vm = $i.VM
        [pscustomobject]@{
            'VM' = $vm.Name; 'Estado' = (Badge $(if ($vm.State -eq 'Running') { 'OK' } elseif ($vm.State -eq 'Off') { 'INFO' } else { 'WARN' }) ([string]$vm.State))
            'Uptime' = (Span $vm.Uptime); 'Geração' = $vm.Generation; 'Versão' = $vm.Version; 'vCPU' = $vm.ProcessorCount
            'Memória' = $(if ($vm.State -eq 'Running') { Bytes $vm.MemoryAssigned } else { Bytes $vm.MemoryStartup })
            'Mem. dinâmica' = $(if ($vm.DynamicMemoryEnabled) { "{0} - {1}" -f (Bytes $vm.MemoryMinimum), (Bytes $vm.MemoryMaximum) } else { 'NÃO' })
            'Heartbeat' = [string]$vm.Heartbeat; 'Sistema convidado' = (PV ([pscustomobject]$i.Kvp) 'OSName')
            'Discos' = (Bytes $i.Bytes); 'Checkpoints' = $i.Snaps.Count; 'Replicação' = [string]$vm.ReplicationState
        }
    }
    [void]$out.Append((Det '03' 'Inventário de VMs' 'Get-VM: estado, recursos e integração. Memória = atribuída (ligadas) ou de inicialização (desligadas).' (Tbl $rows 'VM', 'Estado', 'Uptime', 'Geração', 'Versão', 'vCPU', 'Memória', 'Mem. dinâmica', 'Heartbeat', 'Sistema convidado', 'Discos', 'Checkpoints', 'Replicação' -Raw 'Estado' -Num 'vCPU', 'Memória', 'Discos', 'Checkpoints' -Data) $vms.Count))
    $rows = foreach ($i in $Lista) {
        $k = $i.Kvp
        [pscustomobject]@{ 'VM' = $i.VM.Name; 'Nome no convidado' = $k['FullyQualifiedDomainName']; 'Sistema' = $k['OSName']; 'Versão' = $k['OSVersion']; 'Build' = $k['OSBuildNumber']; 'Integration Services (convidado)' = $k['IntegrationServicesVersion']; 'IPs (Hyper-V)' = ((@($i.Nics | ForEach-Object { $_.IPAddresses }) | Where-Object { $_ -and $_ -notmatch ':' }) -join ', ') }
    }
    [void]$out.Append((Det '03' 'Sistema convidado' 'Dados reportados pelo convidado via KVP (somente VMs ligadas com integration services ativos).' (Tbl $rows 'VM', 'Nome no convidado', 'Sistema', 'Versão', 'Build', 'Integration Services (convidado)', 'IPs (Hyper-V)' -Data) $Lista.Count))
    $rows = foreach ($i in $Lista) {
        $vm = $i.VM; $p = $i.Proc
        [pscustomobject]@{
            'VM' = $vm.Name; 'vCPU' = $vm.ProcessorCount; 'Reserva %' = (PV $p 'Reserve'); 'Limite %' = (PV $p 'Maximum'); 'Peso' = (PV $p 'RelativeWeight')
            'Compat. migração' = $(if ($p) { SN $p.CompatibilityForMigrationEnabled } else { '' }); 'Virtualização aninhada' = $(if ($p) { SN $p.ExposeVirtualizationExtensions } else { '' })
            'Ação início' = [string]$vm.AutomaticStartAction; 'Atraso (s)' = $vm.AutomaticStartDelay; 'Ação parada' = [string]$vm.AutomaticStopAction
            'Checkpoint' = [string]$vm.CheckpointType; 'Secure Boot' = $(if ($i.Fw) { [string]$i.Fw.SecureBoot } else { '-' }); 'Configuração em' = $vm.ConfigurationLocation
        }
    }
    [void]$out.Append((Det '03' 'Processador, ações automáticas e firmware' 'Get-VMProcessor, ações de início/parada, tipo de checkpoint e firmware.' (Tbl $rows 'VM', 'vCPU', 'Reserva %', 'Limite %', 'Peso', 'Compat. migração', 'Virtualização aninhada', 'Ação início', 'Atraso (s)', 'Ação parada', 'Checkpoint', 'Secure Boot', 'Configuração em' -Num 'vCPU', 'Reserva %', 'Limite %', 'Peso', 'Atraso (s)' -Wide 'Configuração em' -Data) $Lista.Count))
    $rows = foreach ($i in $Lista) {
        $vm = $i.VM
        [pscustomobject]@{
            'VM' = $vm.Name; 'Dinâmica' = (SN $vm.DynamicMemoryEnabled); 'Inicialização' = (Bytes $vm.MemoryStartup); 'Mínimo' = $(if ($vm.DynamicMemoryEnabled) { Bytes $vm.MemoryMinimum } else { '' })
            'Máximo' = $(if ($vm.DynamicMemoryEnabled) { Bytes $vm.MemoryMaximum } else { '' }); 'Atribuída' = (Bytes $vm.MemoryAssigned); 'Demanda' = (Bytes $vm.MemoryDemand)
            'Buffer %' = (PV $i.Mem 'Buffer'); 'Prioridade' = (PV $i.Mem 'Priority'); 'Status' = [string]$vm.MemoryStatus
        }
    }
    [void]$out.Append((Det '03' 'Memória' 'Get-VMMemory e uso atual.' (Tbl $rows 'VM', 'Dinâmica', 'Inicialização', 'Mínimo', 'Máximo', 'Atribuída', 'Demanda', 'Buffer %', 'Prioridade', 'Status' -Num 'Inicialização', 'Mínimo', 'Máximo', 'Atribuída', 'Demanda', 'Buffer %', 'Prioridade' -Data) $Lista.Count))
    $rows = foreach ($i in $Lista) { foreach ($s in $i.IS) {
        $ok = "$($s.PrimaryStatusDescription)" -match '^(OK|Ok)$'
        [pscustomobject]@{ 'VM' = $i.VM.Name; 'Serviço' = $s.Name; 'Habilitado' = (SN $s.Enabled); 'Status' = (Badge $(if (-not $s.Enabled) { 'INFO' } elseif ($ok -or -not $s.PrimaryStatusDescription) { 'OK' } else { 'WARN' }) $(if ($s.PrimaryStatusDescription) { [string]$s.PrimaryStatusDescription } else { '-' })); 'Detalhe' = $s.SecondaryStatusDescription }
    } }
    [void]$out.Append((Det '03' 'Integration services' 'Get-VMIntegrationService: status por serviço de integração.' (Tbl $rows 'VM', 'Serviço', 'Habilitado', 'Status', 'Detalhe' -Raw 'Status' -Data) @($rows).Count))
    $rows = foreach ($i in $Lista) { foreach ($d in $i.Dvd) { [pscustomobject]@{ 'VM' = $i.VM.Name; 'Controlador' = ("{0} {1}:{2}" -f $d.ControllerType, $d.ControllerNumber, $d.ControllerLocation); 'Mídia' = $(if ($d.Path) { $d.Path } else { '(vazio)' }) } } }
    [void]$out.Append((Det '03' 'Unidades de DVD' 'ISOs montadas nas VMs.' (Tbl $rows 'VM', 'Controlador', 'Mídia' -Wide 'Mídia' -Data) @($rows).Count))
    [void]$out.Append('</section></div>')

    # --- Armazenamento
    [void]$out.Append('<div class="tabpane" id="t-armazenamento"><section><h2><span class="num">04</span>Armazenamento</h2><p class="lead">Volumes, discos físicos e discos virtuais (VHD/VHDX) das VMs.</p>' + (ChecksArea 'armazenamento'))
    [void]$out.Append('<div class="sub-h">Capacidade</div><div class="grid2">')
    $bv = foreach ($v in $volVis) { $u = 100 - $v.PctLivre; [pscustomobject]@{ N = $v.Nome + $(if ($v.Rotulo) { " ($($v.Rotulo))" } else { '' }); P = $u; C = $(if ($v.PctLivre -lt 10) { 'crit' } elseif ($v.PctLivre -lt 15) { 'warn' } else { 'ok' }); T = ("{0} / {1} · {2}% livre" -f (Bytes ($v.Cap - $v.Livre)), (Bytes $v.Cap), (Num $v.PctLivre 0)) } }
    [void]$out.Append((Card 'Volumes do host' (Bars $bv)))
    $bd = foreach ($d in ($linhasVhd | Where-Object { $_.Max } | Sort-Object Arquivo -Descending)) { [pscustomobject]@{ N = ("{0} · {1}" -f $d.VM, [IO.Path]::GetFileName($d.Caminho)); P = (Pct $d.Arquivo $d.Max); C = 'db'; T = ("{0} de {1} ({2})" -f (Bytes $d.Arquivo), (Bytes $d.Max), $d.Tipo) } }
    [void]$out.Append((Card 'Discos virtuais (ocupado x tamanho máximo)' (Bars $bd)))
    [void]$out.Append('</div>' + (Ferramentas))
    $rows = foreach ($v in $Script:Volumes) {
        $nv = @($linhasVhd | Where-Object { $_.Volume -and $_.Volume.Nome -eq $v.Nome }).Count
        [pscustomobject]@{ 'Volume' = $v.Nome; 'Rótulo' = $v.Rotulo; 'Sistema de arquivos' = $v.FS; 'Tamanho' = (Bytes $v.Cap); 'Usado' = (Bytes ($v.Cap - $v.Livre)); 'Livre' = (Bytes $v.Livre); '% livre' = (Num $v.PctLivre 1); 'Cluster' = $(if ($v.Cluster) { Bytes $v.Cluster } else { '' }); 'Saúde' = $v.Saude; 'Discos de VM' = $nv }
    }
    [void]$out.Append((Det '04' 'Volumes' 'Win32_Volume / Get-Volume.' (Tbl $rows 'Volume', 'Rótulo', 'Sistema de arquivos', 'Tamanho', 'Usado', 'Livre', '% livre', 'Cluster', 'Saúde', 'Discos de VM' -Num 'Tamanho', 'Usado', 'Livre', '% livre', 'Cluster', 'Discos de VM' -Data) $Script:Volumes.Count))
    $rows = foreach ($d in $linhasVhd) {
        [pscustomobject]@{ 'VM' = $d.VM; 'Controlador' = $d.Controlador; 'Arquivo' = $d.Caminho; 'Formato' = $d.Formato; 'Tipo' = $d.Tipo; 'Tamanho máximo' = (Bytes $d.Max); 'Ocupado' = (Bytes $d.Arquivo); 'Fragmentação %' = $d.Frag; 'Níveis' = $d.Niveis; 'Setor (lóg/fís)' = $d.Setor; 'Volume' = $(if ($d.Volume) { $d.Volume.Nome } else { '' }) }
    }
    [void]$out.Append((Det '04' 'Discos virtuais das VMs' 'Get-VMHardDiskDrive + Get-VHD. Níveis = arquivos na cadeia (1 = sem diferenciais).' (Tbl $rows 'VM', 'Controlador', 'Arquivo', 'Formato', 'Tipo', 'Tamanho máximo', 'Ocupado', 'Fragmentação %', 'Níveis', 'Setor (lóg/fís)', 'Volume' -Num 'Tamanho máximo', 'Ocupado', 'Fragmentação %', 'Níveis' -Wide 'Arquivo' -Data) $linhasVhd.Count))
    $rows = foreach ($f in $orfaos) { [pscustomobject]@{ 'Arquivo' = $f.FullName; 'Tamanho' = (Bytes $f.Length); 'Modificado em' = (Dt $f.LastWriteTime) } }
    [void]$out.Append((Det '04' 'Arquivos de disco virtual não anexados' 'Arquivos .vhd/.vhdx/.avhdx nas pastas das VMs e na pasta padrão que não pertencem à cadeia de nenhuma VM. Nada foi alterado; validar antes de remover.' (Tbl $rows 'Arquivo', 'Tamanho', 'Modificado em' -Num 'Tamanho' -Wide 'Arquivo' -Data) $orfaos.Count))
    $rows = foreach ($p in $pdisks) { [pscustomobject]@{ 'ID' = $p.DeviceId; 'Disco' = $p.FriendlyName; 'Mídia' = [string]$p.MediaType; 'Barramento' = [string]$p.BusType; 'Tamanho' = (Bytes $p.Size); 'Saúde' = (Badge $(if ("$($p.HealthStatus)" -eq 'Healthy') { 'OK' } else { 'CRIT' }) ([string]$p.HealthStatus)); 'Status operacional' = (@($p.OperationalStatus) -join ', '); 'Firmware' = $p.FirmwareVersion; 'Série' = ("$($p.SerialNumber)").Trim() } }
    [void]$out.Append((Det '04' 'Discos físicos' 'Get-PhysicalDisk. Em RAID de hardware aparece o disco lógico da controladora.' (Tbl $rows 'ID', 'Disco', 'Mídia', 'Barramento', 'Tamanho', 'Saúde', 'Status operacional', 'Firmware', 'Série' -Raw 'Saúde' -Num 'Tamanho' -Data) $pdisks.Count))
    $rows = foreach ($d in $disks) { [pscustomobject]@{ 'Nº' = $d.Number; 'Disco' = $d.FriendlyName; 'Partição' = [string]$d.PartitionStyle; 'Tamanho' = (Bytes $d.Size); 'Status' = [string]$d.OperationalStatus; 'Saúde' = [string]$d.HealthStatus; 'Offline' = (SN $d.IsOffline); 'Somente leitura' = (SN $d.IsReadOnly); 'Boot/Sistema' = (SN ($d.IsBoot -or $d.IsSystem)) } }
    [void]$out.Append((Det '04' 'Discos (visão do Windows)' 'Get-Disk.' (Tbl $rows 'Nº', 'Disco', 'Partição', 'Tamanho', 'Status', 'Saúde', 'Offline', 'Somente leitura', 'Boot/Sistema' -Num 'Nº', 'Tamanho' -Data) $disks.Count))
    if ($rel.Count) {
        $rows = foreach ($x in $rel) { [pscustomobject]@{ 'Disco' = $x.D.FriendlyName; 'Temperatura (°C)' = $x.R.Temperature; 'Desgaste %' = $x.R.Wear; 'Erros de leitura' = $x.R.ReadErrorsTotal; 'Erros de escrita' = $x.R.WriteErrorsTotal; 'Horas ligado' = $x.R.PowerOnHours; 'Latência máx. leitura (ms)' = $x.R.ReadLatencyMax; 'Latência máx. escrita (ms)' = $x.R.WriteLatencyMax } }
        [void]$out.Append((Det '04' 'Contadores de confiabilidade' 'Get-StorageReliabilityCounter (quando suportado pela controladora).' (Tbl $rows 'Disco', 'Temperatura (°C)', 'Desgaste %', 'Erros de leitura', 'Erros de escrita', 'Horas ligado', 'Latência máx. leitura (ms)', 'Latência máx. escrita (ms)' -Num 'Temperatura (°C)', 'Desgaste %', 'Erros de leitura', 'Erros de escrita', 'Horas ligado', 'Latência máx. leitura (ms)', 'Latência máx. escrita (ms)' -Data) $rel.Count))
    }
    [void]$out.Append('</section></div>')

    # --- Rede
    [void]$out.Append('<div class="tabpane" id="t-rede"><section><h2><span class="num">05</span>Rede</h2><p class="lead">Switches virtuais, placas físicas, teaming, IPs do host e adaptadores das VMs.</p>' + (ChecksArea 'rede') + (Ferramentas))
    $rows = foreach ($s in $switches) { [pscustomobject]@{ 'Switch' = $s.Name; 'Tipo' = [string]$s.SwitchType; 'Uplink' = ((@(PV $s 'NetAdapterInterfaceDescriptions', 'NetAdapterInterfaceDescription')) -join '; '); 'Gerenciamento (host)' = (SN $s.AllowManagementOS); 'SET' = (SN $s.EmbeddedTeamingEnabled); 'SR-IOV' = (SN $s.IovEnabled); 'Modo de banda' = [string]$s.BandwidthReservationMode; 'VMs' = @($Lista | Where-Object { $_.Nics | Where-Object { $_.SwitchName -eq $s.Name } }).Count } }
    [void]$out.Append((Det '05' 'Switches virtuais' 'Get-VMSwitch.' (Tbl $rows 'Switch', 'Tipo', 'Uplink', 'Gerenciamento (host)', 'SET', 'SR-IOV', 'Modo de banda', 'VMs' -Num 'VMs' -Wide 'Uplink' -Data) $switches.Count))
    $rows = foreach ($a in ($nicsAll | Sort-Object Name)) {
        $q = $vmq | Where-Object { $_.Name -eq $a.Name } | Select-Object -First 1
        $dd = ''; try { if ($a.DriverDate) { $dd = ([datetime]$a.DriverDate).ToString('dd/MM/yyyy') } } catch { $dd = [string]$a.DriverDate }
        [pscustomobject]@{ 'Adaptador' = $a.Name; 'Descrição' = $a.InterfaceDescription; 'Status' = (Badge $(if ($a.Status -eq 'Up') { 'OK' } else { 'INFO' }) ([string]$a.Status)); 'Velocidade' = $a.LinkSpeed; 'MAC' = $a.MacAddress; 'Físico' = (SN $a.HardwareInterface); 'Uplink de switch' = (SN ($uplinkDesc -contains $a.InterfaceDescription)); 'VMQ' = $(if ($q) { SN $q.Enabled } else { '' }); 'Driver' = $a.DriverVersion; 'Data do driver' = $dd }
    }
    [void]$out.Append((Det '05' 'Adaptadores de rede do host' 'Get-NetAdapter / Get-NetAdapterVmq.' (Tbl $rows 'Adaptador', 'Descrição', 'Status', 'Velocidade', 'MAC', 'Físico', 'Uplink de switch', 'VMQ', 'Driver', 'Data do driver' -Raw 'Status' -Data) $nicsAll.Count))
    $rows = foreach ($c in $ipcfg) { [pscustomobject]@{ 'Interface' = $c.InterfaceAlias; 'IPv4' = (@($c.IPv4Address | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', '); 'Gateway' = (@($c.IPv4DefaultGateway | ForEach-Object { $_.NextHop }) -join ', '); 'DNS' = (@($c.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | ForEach-Object { $_.ServerAddresses }) -join ', '); 'Perfil' = $(if ($c.NetProfile) { "$($c.NetProfile.Name) ($($c.NetProfile.NetworkCategory))" } else { '' }) } }
    [void]$out.Append((Det '05' 'Configuração IP do host' 'Get-NetIPConfiguration.' (Tbl $rows 'Interface', 'IPv4', 'Gateway', 'DNS', 'Perfil' -Data) $ipcfg.Count))
    if ($teams.Count) {
        $rows = foreach ($t in $teams) { [pscustomobject]@{ 'Team' = $t.Name; 'Modo' = [string]$t.TeamingMode; 'Balanceamento' = [string]$t.LoadBalancingAlgorithm; 'Status' = [string]$t.Status; 'Membros' = (@($t.Members) -join ', ') } }
        [void]$out.Append((Det '05' 'NIC Teaming (LBFO)' 'Get-NetLbfoTeam.' (Tbl $rows 'Team', 'Modo', 'Balanceamento', 'Status', 'Membros' -Data) $teams.Count))
    }
    $rows = foreach ($h in $hostVnic) { [pscustomobject]@{ 'Adaptador' = $h.Name; 'Switch' = $h.SwitchName; 'MAC' = $h.MacAddress; 'Status' = (@($h.Status) -join ', ') } }
    [void]$out.Append((Det '05' 'Adaptadores virtuais do host (ManagementOS)' 'vNICs do sistema operacional do host nos switches virtuais.' (Tbl $rows 'Adaptador', 'Switch', 'MAC', 'Status' -Data) $hostVnic.Count))
    $rows = foreach ($i in $Lista) { foreach ($nic in $i.Nics) {
        $vl = Coletar "VM $($i.VM.Name) - VLAN" { Get-VMNetworkAdapterVlan -VMNetworkAdapter $nic }
        [pscustomobject]@{ 'VM' = $i.VM.Name; 'Adaptador' = $nic.Name; 'Switch' = $(if ($nic.SwitchName) { $nic.SwitchName } else { '(desconectado)' }); 'MAC' = $nic.MacAddress; 'MAC dinâmico' = (SN $nic.DynamicMacAddressEnabled); 'VLAN' = $(if ($vl) { if ("$($vl.OperationMode)" -eq 'Access') { $vl.AccessVlanId } else { [string]$vl.OperationMode } } else { '' }); 'IPs' = (@($nic.IPAddresses) -join ', '); 'Status' = (@($nic.Status) -join ', ') }
    } }
    [void]$out.Append((Det '05' 'Adaptadores de rede das VMs' 'Get-VMNetworkAdapter / Get-VMNetworkAdapterVlan. VLAN 0 = sem tag.' (Tbl $rows 'VM', 'Adaptador', 'Switch', 'MAC', 'MAC dinâmico', 'VLAN', 'IPs', 'Status' -Wide 'IPs' -Data) @($rows).Count))
    [void]$out.Append('</section></div>')

    # --- Checkpoints e Backup
    [void]$out.Append('<div class="tabpane" id="t-protecao"><section><h2><span class="num">06</span>Checkpoints, replicação e backup</h2><p class="lead">Checkpoints existentes, Hyper-V Replica, VSS writers e software de backup.</p>' + (ChecksArea 'protecao') + (Ferramentas))
    $rows = foreach ($s in ($todosSnaps | Sort-Object VMName, CreationTime)) { [pscustomobject]@{ 'VM' = $s.VMName; 'Checkpoint' = $s.Name; 'Tipo' = [string]$s.SnapshotType; 'Criado em' = (Dt $s.CreationTime); 'Idade (dias)' = [math]::Floor(($Script:Agora - $s.CreationTime).TotalDays); 'Pai' = $s.ParentSnapshotName } }
    [void]$out.Append((Det '06' 'Checkpoints' 'Get-VMSnapshot.' (Tbl $rows 'VM', 'Checkpoint', 'Tipo', 'Criado em', 'Idade (dias)', 'Pai' -Num 'Idade (dias)' -Data) $todosSnaps.Count))
    if ($replAll.Count) {
        $rows = foreach ($r in $replAll) { [pscustomobject]@{ 'VM' = $r.VMName; 'Modo' = [string]$r.Mode; 'Estado' = [string]$r.State; 'Saúde' = (Badge $(switch ("$($r.Health)") { 'Normal' { 'OK' } 'Warning' { 'WARN' } 'Critical' { 'CRIT' } default { 'INFO' } }) ([string]$r.Health)); 'Primário' = $r.PrimaryServer; 'Réplica' = $r.ReplicaServer; 'Frequência (s)' = (PV $r 'FrequencySec', 'ReplicationFrequencySec'); 'Última replicação' = (Dt (PV $r 'LastReplicationTime')) } }
        [void]$out.Append((Det '06' 'Hyper-V Replica' 'Get-VMReplication.' (Tbl $rows 'VM', 'Modo', 'Estado', 'Saúde', 'Primário', 'Réplica', 'Frequência (s)', 'Última replicação' -Raw 'Saúde' -Num 'Frequência (s)' -Data) $replAll.Count))
    }
    [void]$out.Append((Det '06' 'VSS writers' 'vssadmin list writers (somente leitura).' (Tbl $writers 'Writer', 'Estado', 'Último erro', 'Status' -Raw 'Status' -Data) $writers.Count))
    $rows = foreach ($s in $bkpSvc) { [pscustomobject]@{ 'Serviço' = $s.Name; 'Descrição' = $s.DisplayName; 'Início' = $s.StartMode; 'Estado' = $s.State } }
    [void]$out.Append((Det '06' 'Software de backup identificado' 'Serviços de backup conhecidos instalados no host.' (Tbl $rows 'Serviço', 'Descrição', 'Início', 'Estado' -Data) $bkpSvc.Count))
    if ($wbs) {
        $kvW = [ordered]@{ 'Último backup com sucesso' = (Dt $wbs.LastSuccessfulBackupTime); 'Último backup' = (Dt $wbs.LastBackupTime); 'Resultado do último' = $wbs.LastBackupResultHR; 'Próximo agendado' = (Dt $wbs.NextBackupTime); 'Versões disponíveis' = $wbs.NumberOfVersions }
        [void]$out.Append((Det '06' 'Windows Server Backup' 'Get-WBSummary.' (KV $kvW) $null))
    }
    [void]$out.Append('</section></div>')

    # --- Eventos
    [void]$out.Append('<div class="tabpane" id="t-eventos"><section><h2><span class="num">07</span>Eventos</h2><p class="lead">Erros e avisos dos últimos ' + $DiasEventos + ' dias, agrupados por origem e ID (desde ' + (Dt $desde) + ').</p>' + (ChecksArea 'eventos') + (Ferramentas))
    $cols = 'Log', 'Origem', 'ID', 'Nível', 'Ocorrências', 'Primeira', 'Última', 'Mensagem (última)'
    $rs = @(Resumir-Eventos $evHv)
    [void]$out.Append((Det '07' 'Logs do Hyper-V (Admin)' 'Críticos, erros e avisos de Microsoft-Windows-Hyper-V-*-Admin.' (Tbl $rs $cols -Raw 'Nível' -Num 'ID', 'Ocorrências' -Wide 'Mensagem (última)' -Data) $rs.Count))
    $rs = @(Resumir-Eventos ($evDisco + $evDesl))
    [void]$out.Append((Det '07' 'Disco, sistema de arquivos e desligamentos inesperados' 'Eventos de storage/NTFS e Kernel-Power 41 / EventLog 6008.' (Tbl $rs $cols -Raw 'Nível' -Num 'ID', 'Ocorrências' -Wide 'Mensagem (última)' -Data) $rs.Count))
    $rs = @(Resumir-Eventos $evSys | Select-Object -First 40)
    [void]$out.Append((Det '07' 'System (críticos e erros)' 'Até 40 grupos mais relevantes.' (Tbl $rs $cols -Raw 'Nível' -Num 'ID', 'Ocorrências' -Wide 'Mensagem (última)' -Data) $rs.Count))
    $rs = @(Resumir-Eventos $evApp | Select-Object -First 40)
    [void]$out.Append((Det '07' 'Application (críticos e erros)' 'Até 40 grupos mais relevantes.' (Tbl $rs $cols -Raw 'Nível' -Num 'ID', 'Ocorrências' -Wide 'Mensagem (última)' -Data) $rs.Count))
    [void]$out.Append('</section></div>')

    # --- Desempenho
    [void]$out.Append('<div class="tabpane" id="t-performance"><section><h2><span class="num">08</span>Desempenho</h2><p class="lead">Amostra de ' + $AmostraSegundos + ' segundos durante a coleta. É um retrato do momento, não substitui monitoramento contínuo.</p>' + (ChecksArea 'performance'))
    [void]$out.Append('<div class="sub-h">Amostra</div><div class="grid2">')
    $bp = @()
    if ($cpuLP) { $bp += [pscustomobject]@{ N = 'Hipervisor (todas as CPUs)'; P = $cpuLP.Media; C = $(if ($cpuLP.Media -gt 85) { 'crit' } elseif ($cpuLP.Media -gt 70) { 'warn' } else { 'ok' }); T = (Num $cpuLP.Media 1) + '%' } }
    if ($cpuRoot) { $bp += [pscustomobject]@{ N = 'Partição raiz (host)'; P = $cpuRoot.Media; C = 'db'; T = (Num $cpuRoot.Media 1) + '%' } }
    foreach ($vm in $vmOn) { $bp += [pscustomobject]@{ N = $vm.Name; P = $vm.CPUUsage; C = 'db'; T = "$($vm.CPUUsage)% de $($vm.ProcessorCount) vCPU" } }
    [void]$out.Append((Card 'CPU' (Bars $bp)))
    $bl = foreach ($l in $latencias) { $m = [math]::Max($l._lr, $l._lw); [pscustomobject]@{ N = $l.Volume; P = [math]::Min([double]100, $m * 2); C = $(if ($m -gt 50) { 'crit' } elseif ($m -gt 20) { 'warn' } else { 'ok' }); T = ("L {0} ms · E {1} ms" -f $l.'Latência leitura (ms)', $l.'Latência escrita (ms)') } }
    [void]$out.Append((Card 'Latência de disco (escala até 50 ms)' (Bars $bl)))
    [void]$out.Append('</div>' + (Ferramentas))
    [void]$out.Append((Det '08' 'Volumes: IOPS e latência' 'Win32_PerfRawData_PerfDisk_LogicalDisk (duas leituras no intervalo da amostra).' (Tbl $latencias 'Volume', 'Leituras/s', 'Escritas/s', 'Latência leitura (ms)', 'Latência escrita (ms)', 'Fila atual' -Num 'Leituras/s', 'Escritas/s', 'Latência leitura (ms)', 'Latência escrita (ms)', 'Fila atual' -Data) $latencias.Count))
    $rows = foreach ($vm in $vms) { [pscustomobject]@{ 'VM' = $vm.Name; 'Estado' = [string]$vm.State; 'CPU %' = $vm.CPUUsage; 'vCPU' = $vm.ProcessorCount; 'Memória atribuída' = (Bytes $vm.MemoryAssigned); 'Demanda' = (Bytes $vm.MemoryDemand); 'Status memória' = [string]$vm.MemoryStatus } }
    [void]$out.Append((Det '08' 'VMs: CPU e memória no momento' 'Get-VM.' (Tbl $rows 'VM', 'Estado', 'CPU %', 'vCPU', 'Memória atribuída', 'Demanda', 'Status memória' -Num 'CPU %', 'vCPU', 'Memória atribuída', 'Demanda' -Data) $vms.Count))
    [void]$out.Append('</section></div>')

    # --- Checklist
    [void]$out.Append('<div class="tabpane" id="t-checklist"><section id="checklist"><h2><span class="num">09</span>Checklist do assessment</h2><p class="lead">Todas as verificações executadas, agrupadas por categoria.</p><div class="tw"><table><thead><tr><th>Indicador</th><th>Valor</th><th>Status</th><th>Observação</th></tr></thead><tbody>')
    foreach ($g in ($Script:Checks | Group-Object Cat | Sort-Object Name)) {
        [void]$out.Append('<tr class="grp"><td colspan="4">' + (Esc $g.Name) + '</td></tr>')
        foreach ($x in $g.Group) { [void]$out.Append('<tr><td>' + (Esc $x.Ind) + '</td><td class="n">' + (Esc $x.Val) + '</td><td>' + (Badge $x.Status) + '</td><td class="obs">' + (Esc $x.Obs) + '</td></tr>') }
    }
    [void]$out.Append('</tbody></table></div></section><section id="coleta"><h2><span class="num">10</span>Observações da coleta</h2><p class="lead">Itens que não puderam ser coletados (recurso não instalado, sem permissão ou não suportado). Nenhuma alteração foi feita no servidor.</p>')
    $notas = @($Script:Notas | Sort-Object 'Seção' -Unique)
    [void]$out.Append($(if ($notas.Count) { Tbl $notas 'Seção', 'Mensagem' -Wide 'Mensagem' } else { '<p class="empty">Todas as coletas foram concluídas sem erro.</p>' }))
    [void]$out.Append('</section></div>')

    # --- Rodapé
    [void]$out.Append('<footer class="foot"><img alt="Solo Network - Rod. João Leopoldo Jacomel, 12.162 - Pinhais/PR - 0800 604 9596 - www.solonetwork.com.br" src="data:image/png;base64,' + $Script:RodapeB64 + '"/><div class="gen">' + (Esc $Script:Versao) + ' · dados coletados em ' + (Dt $Script:Agora) + ' · coleta somente leitura</div></footer></div></body></html>')

    $nomeArq = "Assessment_HyperV_{0}_{1}.html" -f ($hostNome -replace '[\\/:*?"<>|]', '_'), $Script:Agora.ToString('yyyyMMdd_HHmm')
    $arquivo = Join-Path $PastaSaida $nomeArq
    [IO.File]::WriteAllText($arquivo, $out.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    Write-Host ''
    Write-Host ("Crítico: {0} · Atenção: {1} · OK: {2} · Info: {3}" -f $nCrit, $nWarn, $nOk, $nInfo) -ForegroundColor Yellow
    Write-Host "Relatório gerado: $arquivo" -ForegroundColor Green
    if ($Abrir) { Invoke-Item -LiteralPath $arquivo }
}

# =================================================================================================
#  Recursos visuais Solo Network (papel timbrado e estilo do relatório padrão). Não editar.
# =================================================================================================
$Script:Cabecalho = @'
<html lang="pt-BR"><head><meta charset="utf-8"/><meta content="width=device-width, initial-scale=1" name="viewport"/><title>{{TITULO}}</title><link href="https://fonts.googleapis.com/css2?family=Lato:wght@400;700&amp;display=swap" rel="stylesheet"/><style>:root{--red:#D32225;--red-dk:#A72A2B;--ink:#2F2F2E;--ink2:#3C3C3B;--muted:#76766F;--line:#E4E4E1;--soft:#F6F6F4;--bg:#EDEDEA;--ok:#2E7D4F;--ok-bg:#E6F2EA;--warn:#A86400;--warn-bg:#FFF2DB;--crit:#B3121B;--crit-bg:#FBE3E4;--info:#4F6272;--info-bg:#E9EEF2}
*{box-sizing:border-box}
html{background:var(--bg)}
body{margin:0;padding:32px 16px;background:var(--bg);color:var(--ink);font:14px/1.5 "Lato","Segoe UI",Roboto,Helvetica,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.page{position:relative;max-width:1180px;margin:0 auto;background:#fff url(data:image/png;base64,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) no-repeat right 0 top 28px;background-size:190px auto;box-shadow:0 2px 24px rgba(0,0,0,.08);padding:56px 64px 28px 76px;overflow:hidden}
.page::before{content:"";position:absolute;left:0;top:92px;width:30px;height:168px;background:url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADoAAAE+CAYAAADcXQOKAAADLUlEQVR42u3awW2jUBDG8c+R70kHDpflwCEuwR04JTgVrDvYlEAHSzogFYR0gKXlwAnTAa6APSxICDmJLewsvPmP5IMPHH4eZt578zyr61oW4kZG4kaSMs9fZp5/7zJ09uf+Ryxp3XzfSUraT1DklUvQz4r0vYNOXIZ249DLduoqtB9lD753FdqPUdf3JaGjru9rQkdV398FPQaPv7O+/xf0WGOLJUXXyvZYoH102KArl6Hd1zuUFF4CPGboRcFTOL3cSvolKc08f2XhmLaQ9JZ5fph5/p2F8+hPSUnm+UsLB++HBrtyHdrW7lvm+RvXoW38PgXryszoS6xLw7HwswblEvRWUvzR0uPauHMhKbIAlaT1sWXH1QF2ZAW66Hdhl68ktlagD93lxvVLpo0V6KMV6KK9JbRwP7qyAjWT0aUV6J0VqIACBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoUKFCgQIECBQoU6PCorEBTK9C9FWhiAVoGRW4io4mVrhtagJZBkacWoM8WNgxlUOSRBejWwhbwNSjy2HXoQdLGwqZ+FRR55Tr0qbucuAp96ndZF6FfIiVpPvHG8xgUeeLyhGHXNJ7k1AemCH1pkOk5D03p1S0lbY9tBlyBHprjVvjRGukC9EXSczslGBLzkb6i0dAMjhVaSoolRec2mbFDD/o3z4klJZd4NccEfe/A0u/+Za8J3bVZO2dhnwK0rbOkyVqlEcV8YJ11YXuNOM6FvnZgqSYU8xPqrG0giSYc8yZL6ynU2ZCY1XWtzPOXkqqx19lgqIUw8x+GvxtS+i0JqTS8AAAAAElFTkSuQmCC) no-repeat left top;background-size:100% 100%}
a{color:inherit}
h1,h2,h3{font-weight:700;color:var(--ink2);margin:0}
.cover{position:relative;padding:8px 210px 30px 0;border-bottom:1px solid var(--line)}
.kicker{font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:var(--red);font-weight:700}
.cover h1{font-size:30px;line-height:1.2;margin:10px 0 4px;letter-spacing:-.01em}
.client{font-size:16px;color:var(--muted)}
.client b{color:var(--ink2)}
.server{margin:22px 0 0;display:flex;flex-direction:column}
.server .lbl{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted)}
.server .val{font-size:26px;font-weight:700;color:var(--ink2);font-variant-numeric:tabular-nums;word-break:break-all}
.meta{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px 24px;margin-top:22px}
.meta div{border-left:3px solid var(--red);padding:2px 0 2px 10px}
.meta span{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.meta b{display:block;font-size:14px;color:var(--ink2);word-break:break-word}
.toc{display:flex;flex-wrap:wrap;gap:6px 18px;padding:14px 0;border-bottom:1px solid var(--line);font-size:13px}
.toc a{text-decoration:none;color:var(--muted)}
.toc a:hover{color:var(--red)}
.toc a b{color:var(--red);margin-right:6px}
section{margin-top:40px}
h2{display:flex;align-items:baseline;gap:12px;font-size:21px;margin-bottom:16px;padding-bottom:8px;border-bottom:2px solid var(--ink2)}
h2 .num{color:var(--red);font-size:15px;letter-spacing:.06em}
h3{font-size:15px;margin-bottom:12px}
.lead{color:var(--muted);margin:-6px 0 16px;max-width:820px}
.exec{display:grid;grid-template-columns:230px 1fr;gap:24px;align-items:stretch}
.gauge{display:flex;flex-direction:column;align-items:center;justify-content:center;background:var(--soft);border-radius:10px;padding:18px}
.gauge svg{width:150px;height:150px}
.gauge .pct{font-size:30px;font-weight:700;fill:var(--ink2)}
.gauge .cap{font-size:12px;color:var(--muted);text-align:center;margin-top:6px;max-width:170px}
.counts{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.cnt{border-radius:10px;padding:16px 18px;background:var(--soft);display:flex;flex-direction:column;justify-content:space-between;border-top:4px solid var(--info)}
.cnt span{font-size:12px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);font-weight:700}
.cnt b{font-size:40px;line-height:1.1;color:var(--ink2);font-variant-numeric:tabular-nums}
.cnt em{font-style:normal;font-size:12px;color:var(--muted)}
.cnt.crit{border-top-color:var(--crit)}.cnt.crit b{color:var(--crit)}
.cnt.warn{border-top-color:var(--warn)}.cnt.warn b{color:var(--warn)}
.cnt.ok{border-top-color:var(--ok)}.cnt.ok b{color:var(--ok)}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:16px}
.kpi{border:1px solid var(--line);border-radius:10px;padding:14px 16px;background:#fff}
.kpi .k{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.kpi .v{display:block;font-size:24px;font-weight:700;color:var(--ink2);line-height:1.3;font-variant-numeric:tabular-nums;word-break:break-word}
.kpi .s{display:block;font-size:12px;color:var(--muted)}
.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
.card{border:1px solid var(--line);border-radius:10px;padding:18px 20px;background:#fff}
.card+.card{margin-top:0}
.grid2+.card{margin-top:16px}
.hb{display:grid;grid-template-columns:minmax(0,34%) 1fr auto;gap:4px 12px;align-items:center;font-size:13px}
.hb .n{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--ink2)}
.hb .t{color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap;text-align:right}
.bar{height:10px;border-radius:5px;background:var(--line);overflow:hidden}
.bar i{display:block;height:100%;border-radius:5px;background:var(--ink2);font-size:0;line-height:0;color:transparent}
.bar i.ok{background:var(--ok)}.bar i.warn{background:var(--warn)}.bar i.crit{background:var(--crit)}.bar i.db{background:var(--red)}
.empty{color:var(--muted);font-style:italic;font-size:13px}
.cpu svg{width:100%;height:170px;display:block}
.cpu .grid{stroke:var(--line);stroke-width:1}
.cpu .ax{font-size:10px;fill:var(--muted)}
.cpu .l-sql{fill:none;stroke:var(--red);stroke-width:2}
.cpu .l-oth{fill:none;stroke:#9A9A94;stroke-width:1.5;stroke-dasharray:4 3}
.legend{display:flex;gap:18px;font-size:12px;color:var(--muted);margin-top:6px}
.legend span::before{content:"";display:inline-block;width:14px;height:3px;margin-right:6px;vertical-align:middle;background:var(--red)}
.legend span.o::before{background:#9A9A94}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:8px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{background:var(--ink2);color:#fff;text-align:left;font-weight:700;padding:8px 10px;white-space:nowrap;font-size:11.5px;letter-spacing:.02em}
td{padding:7px 10px;border-top:1px solid var(--line);vertical-align:top}
tbody tr:nth-child(even) td{background:#FAFAF8}
td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
table.data td{white-space:nowrap;max-width:520px;overflow:hidden;text-overflow:ellipsis}
table.data td.w{white-space:normal;min-width:260px}
tr.grp td{background:var(--soft)!important;font-weight:700;color:var(--ink2);font-size:12px;letter-spacing:.06em;text-transform:uppercase;border-top:1px solid #d6d6d2}
.b{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:700;letter-spacing:.03em;white-space:nowrap;background:var(--info-bg);color:var(--info)}
.b.ok{background:var(--ok-bg);color:var(--ok)}
.b.warn{background:var(--warn-bg);color:var(--warn)}
.b.crit{background:var(--crit-bg);color:var(--crit)}
td.obs{color:var(--muted)}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px}
.chips a{text-decoration:none;font-size:12px;border:1px solid var(--line);border-radius:14px;padding:3px 10px;color:var(--ink2);background:#fff}
.chips a:hover{border-color:var(--red);color:var(--red)}
.chips a b{color:var(--red);margin-right:4px}
details{border:1px solid var(--line);border-radius:10px;margin-bottom:10px;background:#fff}
details[open]{padding-bottom:14px}
summary{cursor:pointer;list-style:none;padding:12px 16px;display:flex;align-items:center;gap:12px;font-weight:700;color:var(--ink2)}
summary::-webkit-details-marker{display:none}
summary::before{content:"+";width:18px;height:18px;border-radius:50%;background:var(--soft);color:var(--red);display:inline-flex;align-items:center;justify-content:center;font-size:14px;flex:none}
details[open]>summary::before{content:"\2212"}
summary .g{color:var(--red);font-size:12px;letter-spacing:.05em}
summary .c{margin-left:auto;font-weight:400;font-size:12px;color:var(--muted)}
details>.desc,details>.tw,details>.note{margin:0 16px}
details>.desc{color:var(--muted);font-size:13px;margin-bottom:10px}
.note{background:var(--soft);border-left:3px solid var(--info);padding:10px 14px;border-radius:4px;font-size:13px}
.note.warn{border-left-color:var(--warn)}
.foot{margin-top:48px;padding-top:18px}
.foot img{display:block;width:100%;max-width:1106px;height:auto}
.gen{margin-top:10px;font-size:11px;color:var(--muted);text-align:right}
@media (max-width:900px){.page{padding:36px 20px 20px 34px;background-size:120px auto}.page::before{width:18px;height:110px;top:60px}.cover{padding-right:0}.meta,.counts,.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.exec,.grid2{grid-template-columns:1fr}.cover h1{font-size:24px}}
@media print{html,body{background:#fff}body{padding:0;font-size:11px}.page{box-shadow:none;max-width:none;padding:10mm 8mm 10mm 14mm}.toc,.chips{display:none}details{break-inside:auto}details>summary::before{display:none}section{break-inside:auto}h2{break-after:avoid}.card,.kpi,.cnt,tr{break-inside:avoid}table.data td{white-space:normal;max-width:none}.tw{overflow:visible;border:0}}
@page{size:A4;margin:10mm}
.kv tbody th{vertical-align:top;padding:7px 10px;background:#FAFAF8;color:var(--ink2);font-weight:700;font-size:12px;letter-spacing:0;border-top:1px solid var(--line);width:1%;white-space:nowrap}
table.kv td{white-space:normal}
.sub{margin:-6px 0 10px;font-size:12px;color:var(--muted)}
.cpu svg{height:auto}
p.empty{margin:0}
details>.note{margin-top:10px}
.tools{display:flex;gap:8px;margin:-4px 0 12px}
.tools button{font:inherit;font-size:12px;border:1px solid var(--line);background:#fff;color:var(--ink2);border-radius:6px;padding:4px 12px;cursor:pointer}
.tools button:hover{border-color:var(--red);color:var(--red)}
@media print{.tools{display:none}}
table.data td.w{min-width:460px}</style><script>function tg(o){document.querySelectorAll("details").forEach(function(d){d.open=o})}function ab(){var e=document.getElementById(location.hash.substring(1));if(e){if(e.tagName=="DETAILS"){e.open=true}}}window.addEventListener("hashchange",ab);window.addEventListener("load",ab);window.addEventListener("beforeprint",function(){tg(true)});</script><style>
.tabs{position:sticky;top:0;z-index:5;display:flex;flex-wrap:wrap;gap:0 2px;background:#fff;border-bottom:1px solid var(--line);margin:0 -8px;padding:0 8px}
.tabs button{font:inherit;font-size:12.5px;color:var(--muted);background:none;border:0;border-bottom:3px solid transparent;padding:12px 7px 10px;cursor:pointer;white-space:nowrap}
.tabs button:hover{color:var(--ink2)}
.tabs button.on{color:var(--red);border-bottom-color:var(--red);font-weight:700}
.tabs button .n{display:inline-block;min-width:17px;height:17px;line-height:17px;border-radius:9px;font-size:10px;font-weight:700;text-align:center;margin-left:4px;padding:0 4px}
.tabs button .n.crit{background:var(--crit-bg);color:var(--crit)} .tabs button .n.warn{background:var(--warn-bg);color:var(--warn)}
.tabpane{display:none}
.tabpane.on{display:block}
.tabpane>section:first-child{margin-top:28px}
.sub-h{font-size:15px;font-weight:700;color:var(--ink2);margin:26px 0 10px}
@media print{.tabs{display:none}.tabpane{display:block!important;break-before:page}.tabpane:first-of-type{break-before:auto}}
</style><script>function tab(id){document.querySelectorAll('.tabpane').forEach(function(p){p.classList.toggle('on',p.id=='t-'+id)});document.querySelectorAll('.tabs button').forEach(function(b){b.classList.toggle('on',b.getAttribute('data-t')==id)});if(history.replaceState){history.replaceState(null,'','#'+id)}window.scrollTo(0,document.querySelector('.tabs').offsetTop)}window.addEventListener('load',function(){var h=location.hash.substring(1);tab(document.getElementById('t-'+h)?h:'geral')});window.addEventListener('beforeprint',function(){document.querySelectorAll('details').forEach(function(d){d.open=true})});</script></head><body><div class="page">
'@

$Script:RodapeB64 = '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'

Main
