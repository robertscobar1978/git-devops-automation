<#
.SYNOPSIS
    Executa o SQL Server Assessment v11 (Solo Network) e grava o relatório HTML.

.DESCRIPTION
    Lê o script SQL_Server_Assessment_Profissional_v11.sql, ativa @GerarHTML/@SomenteHTML,
    executa na instância informada e salva o relatório em UTF-8, pronto para abrir no navegador.
    V11: relatório organizado em guias (Visão Geral, Atenção, Databases, ... Alta Disponibilidade,
    Checklist) e coleta de Log Shipping (configuração, jobs, atrasos x limites e erros).
    Não depende do módulo SqlServer (usa System.Data.SqlClient do .NET).

.EXAMPLE
    .\Gerar-RelatorioAssessment.ps1 -Instancia "SRV-SQL01\PRODUCAO" -Cliente "TCE-PR" -Abrir

.EXAMPLE
    .\Gerar-RelatorioAssessment.ps1 -Instancia "10.0.0.15,1433" -Usuario "usr_assessment" -Cliente "Cliente X" -ConfiarCertificado

.EXAMPLE
    # Instância muito grande ou horário de pico: pula a coleta de fragmentação de índices
    .\Gerar-RelatorioAssessment.ps1 -Instancia "SRV-SQL01" -Cliente "Cliente X" -SemFragmentacao

.EXAMPLE
    # Avalia fragmentação só de índices com 5.000 páginas ou mais (~40 MB)
    .\Gerar-RelatorioAssessment.ps1 -Instancia "SRV-SQL01" -FragMinPaginas 5000
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Instancia,

    [string]$Cliente = '',

    [string]$Script = (Join-Path $PSScriptRoot 'SQL_Server_Assessment_Profissional_v11.sql'),

    [string]$PastaSaida = (Get-Location).Path,

    # Login SQL (se omitido, usa autenticação Windows). A senha é solicitada de forma segura.
    [string]$Usuario,

    # Criptografa a conexão aceitando certificado autoassinado.
    [switch]$ConfiarCertificado,

    # Abre o relatório no navegador ao final.
    [switch]$Abrir,

    # 0 = sem limite.
    [int]$TimeoutSegundos = 0,

    # V10: não coleta fragmentação de índices (sys.dm_db_index_physical_stats).
    [switch]$SemFragmentacao,

    # V10: tamanho mínimo (em páginas de 8 KB) dos índices avaliados na fragmentação. 0 = padrão do script (1000).
    [int]$FragMinPaginas = 0
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Script)) { throw "Script não encontrado: $Script" }
if (-not (Test-Path -LiteralPath $PastaSaida)) { New-Item -ItemType Directory -Path $PastaSaida | Out-Null }

# --- Ajusta os parâmetros do script -------------------------------------------------------------
$sql = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $Script), [Text.Encoding]::UTF8)
$sql = [regex]::Replace($sql, 'DECLARE @GerarHTML bit = \d;', 'DECLARE @GerarHTML bit = 1;')
$sql = [regex]::Replace($sql, 'DECLARE @SomenteHTML bit = \d;', 'DECLARE @SomenteHTML bit = 1;')
if ($SemFragmentacao) {
    $sql = [regex]::Replace($sql, 'DECLARE @ColetarFragmentacao bit = \d;', 'DECLARE @ColetarFragmentacao bit = 0;')
}
if ($FragMinPaginas -gt 0) {
    $sql = [regex]::Replace($sql, 'DECLARE @FragMinPaginas int = \d+;', "DECLARE @FragMinPaginas int = $FragMinPaginas;")
}
if ($Cliente) {
    $clienteSql = $Cliente.Replace("'", "''")
    $sql = [regex]::Replace($sql, "DECLARE @Cliente nvarchar\(200\) = N'[^']*';",
        [Text.RegularExpressions.MatchEvaluator]{ param($m) "DECLARE @Cliente nvarchar(200) = N'$clienteSql';" })
}

$lotes = [regex]::Split($sql, '^\s*GO\s*$', [Text.RegularExpressions.RegexOptions]'Multiline, IgnoreCase') |
         Where-Object { $_.Trim() -ne '' }

# --- Conexão ------------------------------------------------------------------------------------
$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$csb['Data Source']      = $Instancia
$csb['Initial Catalog']  = 'master'
$csb['Application Name'] = 'Solo Network - SQL Server Assessment v11'
$csb['Connect Timeout']  = 15
if ($Usuario) {
    $cred = Get-Credential -UserName $Usuario -Message "Senha do login SQL '$Usuario' em $Instancia"
    $csb['User ID']  = $Usuario
    $csb['Password'] = $cred.GetNetworkCredential().Password
} else {
    $csb['Integrated Security'] = $true
}
if ($ConfiarCertificado) {
    $csb['Encrypt'] = $true
    $csb['TrustServerCertificate'] = $true
}

$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
# Mensagens (PRINT) e erros não fatais chegam como InfoMessage e não interrompem a leitura.
$conn.FireInfoMessageEventOnUserErrors = $true
$conn.add_InfoMessage({
    param($s, $e)
    foreach ($er in $e.Errors) {
        if ($er.Class -gt 10) { Write-Warning ("[SQL {0}] {1}" -f $er.Number, $er.Message) }
        else { Write-Host $er.Message -ForegroundColor DarkGray }
    }
})

$html   = $null
$partes = New-Object System.Collections.Generic.List[string]

Write-Host "Conectando em $Instancia ..." -ForegroundColor Cyan
$conn.Open()
try {
    foreach ($lote in $lotes) {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $lote
        $cmd.CommandTimeout = $TimeoutSegundos
        $r = $cmd.ExecuteReader()
        try {
            do {
                $iXml = -1; $iTxt = -1
                for ($i = 0; $i -lt $r.FieldCount; $i++) {
                    switch ($r.GetName($i)) {
                        'Relatorio_HTML'       { $iXml = $i }
                        'Relatorio_HTML_Texto' { $iTxt = $i }
                    }
                }
                while ($r.Read()) {
                    if ($iXml -ge 0 -and -not $r.IsDBNull($iXml)) { $html = [string]$r.GetValue($iXml) }
                    elseif ($iTxt -ge 0 -and -not $r.IsDBNull($iTxt)) { $partes.Add([string]$r.GetValue($iTxt)) }
                }
            } while ($r.NextResult())
        }
        finally { $r.Close() }
    }
}
finally { $conn.Close() }

if (-not $html -and $partes.Count -gt 0) { $html = -join $partes }
if (-not $html) { throw 'O script não retornou o relatório HTML. Verifique as mensagens acima.' }

# --- Grava o arquivo ----------------------------------------------------------------------------
$nomeInstancia = $Instancia -replace '[\\/:*?"<>|,]', '_'
$arquivo = Join-Path $PastaSaida ("Assessment_{0}_{1}.html" -f $nomeInstancia, (Get-Date -Format 'yyyyMMdd_HHmm'))
[IO.File]::WriteAllText($arquivo, "<!DOCTYPE html>`r`n" + $html, (New-Object System.Text.UTF8Encoding($true)))

Write-Host "Relatório gerado: $arquivo" -ForegroundColor Green
if ($Abrir) { Invoke-Item -LiteralPath $arquivo }
