# Assessment de Hyper-V (Solo Network)

`Gerar-AssessmentHyperV.ps1` roda no próprio host Hyper-V (Windows Server 2016/2019/2022) e grava um HTML
no padrão Solo (papel timbrado do `Papel Solo.docx` e layout do assessment SQL Server).

**Somente leitura:** só consulta (cmdlets Get-*, CIM, logs de eventos, `vssadmin list writers`, `powercfg /getactivescheme`,
`w32tm /query /source`). Não altera nada no host nem nas VMs. Usa apenas módulos nativos do Windows Server.

## Como executar

1. Copie o `.ps1` para o host (ex.: `C:\Temp`).
2. Abra o PowerShell **como Administrador**.
3. Execute:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
cd C:\Temp
.\Gerar-AssessmentHyperV.ps1 -Cliente "Nome do Cliente" -Abrir
```

Parâmetros opcionais: `-PastaSaida C:\Temp`, `-DiasEventos 30` (janela de eventos), `-AmostraSegundos 15` (amostra de desempenho).
A coleta leva cerca de 1 minuto. O arquivo sai como `Assessment_HyperV_<HOST>_<data>.html`.

## O que o relatório cobre

| Guia | Conteúdo |
|---|---|
| Visão Geral | Resumo executivo, hardware, SO, ciclo de vida, atualizações, reinício pendente, energia, hora, antivírus, configuração do Hyper-V, serviços |
| Atenção | Todos os itens CRÍTICO/ATENÇÃO |
| Máquinas Virtuais | Estado, heartbeat, versão, vCPU, memória, integration services, SO convidado (KVP), ações automáticas, tipo de checkpoint, Secure Boot, ISO |
| Armazenamento | Volumes e espaço livre, discos físicos, VHD/VHDX (tipo, cadeia, fragmentação), crescimento potencial dos discos dinâmicos, arquivos de disco não anexados |
| Rede | Switches virtuais, uplinks, VMQ (Broadcom 1 Gb), teaming, drivers, IPs do host, adaptadores e VLANs das VMs |
| Checkpoints e Backup | Checkpoints e idade, .avhdx órfão, Hyper-V Replica, VSS writers, software de backup, Windows Server Backup |
| Eventos | Logs Hyper-V Admin, erros de disco/NTFS/controladora, desligamentos inesperados, System e Application |
| Desempenho | CPU do hipervisor e da partição raiz, latência/IOPS por volume, CPU e memória das VMs |
| Checklist | Todas as verificações + observações da coleta |

## Manutenção

O CSS e as imagens do papel Solo estão embutidos no fim do script (`$Script:Cabecalho` e `$Script:RodapeB64`),
copiados de `Assessments/gerador/template/`. O script é salvo em UTF-8 com BOM (necessário para acentos no Windows PowerShell 5.1).
