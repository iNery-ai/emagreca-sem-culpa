# Script de geracao da Planilha Profissional de Gestao de Trafego Pago
# Executado via PowerShell para automacao do Microsoft Excel COM

$ErrorActionPreference = "Stop"

# C# helper compiled on the fly to bypass PowerShell's dynamic COM binder caching limitations
$source = @"
using System;
public class ExcelHelper {
    public static void SetValue(object cell, object val) {
        ((dynamic)cell).Value2 = val;
    }
}
"@
try {
    Add-Type -TypeDefinition $source -ReferencedAssemblies "System.Core", "Microsoft.CSharp" -ErrorAction SilentlyContinue
} catch {}


# Funcao para converter Hex Color para o formato BGR inteiro do Excel COM
function Get-ExcelColor {
    param ($Hex)
    $clean = $Hex.Replace("#", "")
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    return ($b * 65536) + ($g * 256) + $r
}

# Cores do Design
$colorHeaderBg = Get-ExcelColor "#1F2937"
$colorCardBg = Get-ExcelColor "#F3F4F6"
$colorZebraBg = Get-ExcelColor "#F9FAFB"
$colorBorder = Get-ExcelColor "#E5E7EB"

# Alertas
$colorSuccessBg = Get-ExcelColor "#D1FAE5"
$colorSuccessText = Get-ExcelColor "#065F46"
$colorWarningBg = Get-ExcelColor "#FEF3C7"
$colorWarningText = Get-ExcelColor "#92400E"
$colorCriticalBg = Get-ExcelColor "#FEE2E2"
$colorCriticalText = Get-ExcelColor "#991B1B"

# Formatos
$fmtCurrency = 'R$ #,##0.00'
$fmtInteger = '#,##0'
$fmtPercent = '0.0%'
$fmtDecimal = '0.00'
$fmtDate = 'dd/mm/yyyy'

Write-Host "Inicializando Excel COM..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Add()
    $sheets = $workbook.Sheets
    while ($sheets.Count -lt 7) {
        $sheets.Add() | Out-Null
    }
    
    $sheet1 = $sheets.Item(1); $sheet1.Name = "Dashboard Executivo"
    $sheet2 = $sheets.Item(2); $sheet2.Name = "Controle Diario"
    $sheet3 = $sheets.Item(3); $sheet3.Name = "Controle de Criativos"
    $sheet4 = $sheets.Item(4); $sheet4.Name = "Analise de Funil"
    $sheet5 = $sheets.Item(5); $sheet5.Name = "Historico de Escala"
    $sheet6 = $sheets.Item(6); $sheet6.Name = "Registro de Vendas"
    $sheet7 = $sheets.Item(7); $sheet7.Name = "Insights Automaticos"
    
    function Set-Cell {
        param ($Sheet, $Row, $Col, $Val, $Format=$null, $Bold=$false, $Align=$null, $Color=$null, $FontColor=$null, $Size=$null)
        try {
            $cell = $Sheet.Cells.Item($Row, $Col)
            if ($Val -is [string]) {
                if ($Val.StartsWith("=")) {
                    $cell.Formula = $Val
                } else {
                    [ExcelHelper]::SetValue($cell, $Val)
                }
            } else {
                [ExcelHelper]::SetValue($cell, $Val)
            }
            
            if ($Format) { $cell.NumberFormat = $Format }
            if ($Bold) { $cell.Font.Bold = $true }
            if ($Size) { $cell.Font.Size = $Size }
            if ($Align -eq "center") { $cell.HorizontalAlignment = -4108 }
            elseif ($Align -eq "right") { $cell.HorizontalAlignment = -4152 }
            elseif ($Align -eq "left") { $cell.HorizontalAlignment = -4131 }
            if ($Color) { $cell.Interior.Color = $Color }
            if ($FontColor) { $cell.Font.Color = $FontColor }
            $cell.Font.Name = "Segoe UI"
        } catch {
            Write-Host "ERRO EM Set-Cell: Row=$Row, Col=$Col, Format=$Format"
            if ($null -eq $Val) {
                Write-Host "Val e nulo"
            } else {
                Write-Host "Val=$Val (Type=$($Val.GetType().FullName))"
            }
            throw $_
        }
    }
    
    function Set-Borders {
        param ($Sheet, $RangeAddress)
        $range = $Sheet.Range($RangeAddress)
        $range.Borders.LineStyle = 1
        $range.Borders.Weight = 2
        $range.Borders.Color = $colorBorder
    }
    
    function Format-TableHeader {
        param ($Sheet, $RangeAddress)
        $range = $Sheet.Range($RangeAddress)
        $range.Font.Name = "Segoe UI"
        $range.Font.Bold = $true
        $range.Font.Size = 11
        $range.Font.Color = 16777215
        $range.Interior.Color = $colorHeaderBg
        $range.HorizontalAlignment = -4108
        $range.VerticalAlignment = -4108
        $range.RowHeight = 26
    }
    
    function Format-TitleHeader {
        param ($Sheet, $RangeAddress, $Title)
        $range = $Sheet.Range($RangeAddress)
        $range.Merge() | Out-Null
        [ExcelHelper]::SetValue($range, $Title)
        $range.Font.Name = "Segoe UI"
        $range.Font.Bold = $true
        $range.Font.Size = 14
        $range.Font.Color = 16777215
        $range.Interior.Color = $colorHeaderBg
        $range.HorizontalAlignment = -4108
        $range.VerticalAlignment = -4108
        $range.RowHeight = 35
    }
    
    $sep = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator

    # =========================================================================
    # ABA 2: CONTROLE DIARIO
    # =========================================================================
    Write-Host "Configurando Controle Diario..."
    Format-TitleHeader $sheet2 "A2:P2" "CONTROLE DIARIO DE TRAFEGO PAGO (PREENCHIMENTO DIARIO)"
    
    $colsDiario = @("Data", "Valor Investido", "Impressoes", "Alcance", "Cliques no Link", "Visualizacoes da Pagina", 
                    "Finalizacoes de Compra", "Compras", "Faturamento", "CPC", "CTR", "CPA", "ROAS", 
                    "Taxa Landing -> Checkout", "Taxa Checkout -> Compra", "Observacoes")
    for ($i = 0; $i -lt $colsDiario.Count; $i++) {
        Set-Cell $sheet2 4 ($i+1) $colsDiario[$i]
    }
    Format-TableHeader $sheet2 "A4:P4"
    
    $diarioData = @(
        @("10/06/2026", 50.00,  5000,  4200,  100, 75,  12, 2, 94.00,  "Campanha 1 ativa"),
        @("11/06/2026", 60.00,  6000,  5100,  120, 90,  15, 3, 141.00, "Subiu orcamento"),
        @("12/06/2026", 75.00,  7500,  6300,  130, 95,  18, 2, 94.00,  "CPA subiu um pouco"),
        @("13/06/2026", 75.00,  7800,  6500,  140, 110, 8,  1, 47.00,  "Queda na conversao"),
        @("14/06/2026", 75.00,  8000,  6800,  160, 125, 22, 4, 188.00, "Recuperacao de boletos"),
        @("15/06/2026", 100.00, 10000, 8500,  210, 160, 30, 5, 235.00, "Novo criativo adicionado"),
        @("16/06/2026", 100.00, 10500, 8900,  220, 170, 28, 6, 282.00, "Melhor dia do mes"),
        @("17/06/2026", 120.00, 12000, 9900,  240, 185, 25, 3, 141.00, "Instabilidade na tarde"),
        @("18/06/2026", 120.00, 12500, 10200, 250, 190, 32, 5, 235.00, "Escalando criativo B"),
        @("19/06/2026", 150.00, 15000, 12000, 310, 240, 45, 8, 376.00, "Excelente performance")
    )
    
    for ($row = 5; $row -le 100; $row++) {
        $isSample = ($row -le 14)
        if ($isSample) {
            $data = $diarioData[$row - 5]
            Set-Cell $sheet2 $row 1  $data[0] -Format $fmtDate -Align "center"
            Set-Cell $sheet2 $row 2  $data[1] -Format $fmtCurrency
            Set-Cell $sheet2 $row 3  $data[2] -Format $fmtInteger
            Set-Cell $sheet2 $row 4  $data[3] -Format $fmtInteger
            Set-Cell $sheet2 $row 5  $data[4] -Format $fmtInteger
            Set-Cell $sheet2 $row 6  $data[5] -Format $fmtInteger
            Set-Cell $sheet2 $row 7  $data[6] -Format $fmtInteger
            Set-Cell $sheet2 $row 8  $data[7] -Format $fmtInteger
            Set-Cell $sheet2 $row 9  $data[8] -Format $fmtCurrency
            Set-Cell $sheet2 $row 16 $data[9] -Align "left"
        } else {
            Set-Cell $sheet2 $row 1  "" -Format $fmtDate -Align "center"
            Set-Cell $sheet2 $row 2  "" -Format $fmtCurrency
            Set-Cell $sheet2 $row 3  "" -Format $fmtInteger
            Set-Cell $sheet2 $row 4  "" -Format $fmtInteger
            Set-Cell $sheet2 $row 5  "" -Format $fmtInteger
            Set-Cell $sheet2 $row 6  "" -Format $fmtInteger
            Set-Cell $sheet2 $row 7  "" -Format $fmtInteger
            Set-Cell $sheet2 $row 8  "" -Format $fmtInteger
            Set-Cell $sheet2 $row 9  "" -Format $fmtCurrency
            Set-Cell $sheet2 $row 16 "" -Align "left"
        }
        
        Set-Cell $sheet2 $row 10 ('=IF(A' + $row + '="","",IF(E' + $row + '>0, B' + $row + '/E' + $row + ', 0))') -Format $fmtCurrency
        Set-Cell $sheet2 $row 11 ('=IF(A' + $row + '="","",IF(C' + $row + '>0, E' + $row + '/C' + $row + ', 0))') -Format $fmtPercent
        Set-Cell $sheet2 $row 12 ('=IF(A' + $row + '="","",IF(H' + $row + '>0, B' + $row + '/H' + $row + ', 0))') -Format $fmtCurrency
        Set-Cell $sheet2 $row 13 ('=IF(A' + $row + '="","",IF(B' + $row + '>0, I' + $row + '/B' + $row + ', 0))') -Format $fmtDecimal
        Set-Cell $sheet2 $row 14 ('=IF(A' + $row + '="","",IF(F' + $row + '>0, G' + $row + '/F' + $row + ', 0))') -Format $fmtPercent
        Set-Cell $sheet2 $row 15 ('=IF(A' + $row + '="","",IF(G' + $row + '>0, H' + $row + '/G' + $row + ', 0))') -Format $fmtPercent
        
        if ($row % 2 -eq 0) {
            $sheet2.Range("A$row:P$row").Interior.Color = $colorZebraBg
        }
        Set-Borders $sheet2 "A$row:P$row"
    }
    
    $roasRange = $sheet2.Range("M5:M100")
    $roasRange.FormatConditions.Delete()
    
    $cfRoasGreen = $roasRange.FormatConditions.Add(2, $null, '=M5>=2')
    $cfRoasGreen.Interior.Color = $colorSuccessBg
    $cfRoasGreen.Font.Color = $colorSuccessText
    
    $cfRoasYellow = $roasRange.FormatConditions.Add(2, $null, '=E(M5>=1;M5<2)')
    $cfRoasYellow.Interior.Color = $colorWarningBg
    $cfRoasYellow.Font.Color = $colorWarningText
    
    $cfRoasRed = $roasRange.FormatConditions.Add(2, $null, '=E(M5>0;M5<1)')
    $cfRoasRed.Interior.Color = $colorCriticalBg
    $cfRoasRed.Font.Color = $colorCriticalText
    
    $cpaRange = $sheet2.Range("L5:L100")
    $cpaRange.FormatConditions.Delete()
    
    $cfCpaGrow = $cpaRange.FormatConditions.Add(2, $null, '=E(L5>L4;É.NÚMERO(L4))')
    $cfCpaGrow.Interior.Color = $colorWarningBg
    $cfCpaGrow.Font.Color = $colorWarningText
    
    $cfCpaAvg = $cpaRange.FormatConditions.Add(2, $null, '=L5>MÉDIA(L$5:L$100)')
    $cfCpaAvg.Interior.Color = $colorCriticalBg
    $cfCpaAvg.Font.Color = $colorCriticalText

    $sheet2.Activate()
    $sheet2.Range("A4:P4").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # =========================================================================
    # ABA 6: REGISTRO DE VENDAS
    # =========================================================================
    Write-Host "Configurando Registro de Vendas..."
    Format-TitleHeader $sheet6 "A2:F2" "REGISTRO DE VENDAS"
    
    $colsVendas = @("Data", "Valor da Venda", "Produto", "Criativo Responsavel", "Campanha", "Observacoes")
    for ($i = 0; $i -lt $colsVendas.Count; $i++) {
        Set-Cell $sheet6 4 ($i+1) $colsVendas[$i]
    }
    Format-TableHeader $sheet6 "A4:F4"
    
    Set-Cell $sheet6 2 8 "Faturamento Total" -Bold $true -Align "center" -Color $colorCardBg
    Set-Cell $sheet6 3 8 "=SUM(B$5:B$1000)" -Format $fmtCurrency -Bold $true -Align "center"
    Set-Borders $sheet6 "H2:H3"
    
    Set-Cell $sheet6 2 9 "Qtd Vendas" -Bold $true -Align "center" -Color $colorCardBg
    Set-Cell $sheet6 3 9 "=COUNT(B$5:B$1000)" -Format $fmtInteger -Bold $true -Align "center"
    Set-Borders $sheet6 "I2:I3"
    
    Set-Cell $sheet6 2 10 "Ticket Medio" -Bold $true -Align "center" -Color $colorCardBg
    Set-Cell $sheet6 3 10 "=IF(I3>0, H3/I3, 0)" -Format $fmtCurrency -Bold $true -Align "center"
    Set-Borders $sheet6 "J2:J3"
    
    $vendasData = @(
        @("10/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("10/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("11/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("11/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("11/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("12/06/2026", 47.00, "Guia Low Ticket", "Criativo C - Video Unboxing", "Campanha Validacao"),
        @("12/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("13/06/2026", 47.00, "Guia Low Ticket", "Criativo D - Imagem Depoimento", "Campanha Validacao"),
        @("14/06/2026", 47.00, "Guia Low Ticket", "Criativo E - Video Promocional", "Campanha Validacao"),
        @("14/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("14/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("14/06/2026", 47.00, "Guia Low Ticket", "Criativo E - Video Promocional", "Campanha Validacao"),
        @("15/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("15/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("15/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("15/06/2026", 47.00, "Guia Low Ticket", "Criativo C - Video Unboxing", "Campanha Validacao"),
        @("15/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("16/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("16/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("16/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("16/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("16/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("16/06/2026", 47.00, "Guia Low Ticket", "Criativo E - Video Promocional", "Campanha Validacao"),
        @("17/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("17/06/2026", 47.00, "Guia Low Ticket", "Criativo C - Video Unboxing", "Campanha Validacao"),
        @("17/06/2026", 47.00, "Guia Low Ticket", "Criativo E - Video Promocional", "Campanha Validacao"),
        @("18/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("18/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("18/06/2026", 47.00, "Guia Low Ticket", "Criativo C - Video Unboxing", "Campanha Validacao"),
        @("18/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("18/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo A - Imagem Oferta", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo B - Video Gancho 1", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo E - Video Promocional", "Campanha Validacao"),
        @("19/06/2026", 47.00, "Guia Low Ticket", "Criativo E - Video Promocional", "Campanha Validacao")
    )
    
    for ($row = 5; $row -le 1000; $row++) {
        $isSample = ($row -le (4 + $vendasData.Count))
        if ($isSample) {
            $data = $vendasData[$row - 5]
            Set-Cell $sheet6 $row 1 $data[0] -Format $fmtDate -Align "center"
            Set-Cell $sheet6 $row 2 $data[1] -Format $fmtCurrency
            Set-Cell $sheet6 $row 3 $data[2] -Align "left"
            Set-Cell $sheet6 $row 4 $data[3] -Align "left"
            Set-Cell $sheet6 $row 5 $data[4] -Align "left"
        } else {
            Set-Cell $sheet6 $row 1 "" -Format $fmtDate -Align "center"
            Set-Cell $sheet6 $row 2 "" -Format $fmtCurrency
            Set-Cell $sheet6 $row 3 "" -Align "left"
            Set-Cell $sheet6 $row 4 "" -Align "left"
            Set-Cell $sheet6 $row 5 "" -Align "left"
        }
        if ($row % 2 -eq 0) {
            $sheet6.Range("A$row:F$row").Interior.Color = $colorZebraBg
        }
        Set-Borders $sheet6 "A$row:F$row"
    }
    
    $sheet6.Activate()
    $sheet6.Range("A4:F4").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # =========================================================================
    # ABA 3: CONTROLE DE CRIATIVOS
    # =========================================================================
    Write-Host "Configurando Controle de Criativos..."
    Format-TitleHeader $sheet3 "A2:L2" "CONTROLE DE CRIATIVOS"
    
    $colsCriativos = @("Nome do Criativo", "Data de Inicio", "Status", "Valor Investido", "Impressoes", "Cliques", 
                       "Finalizacoes de Compra", "Compras", "CPA", "ROAS", "Classificacao", "Observacoes")
    for ($i = 0; $i -lt $colsCriativos.Count; $i++) {
        Set-Cell $sheet3 4 ($i+1) $colsCriativos[$i]
    }
    Format-TableHeader $sheet3 "A4:L4"
    
    $criativosData = @(
        @("Criativo A - Imagem Oferta",    "10/06/2026", "Ativo",    300.00, 30000, 600,  60, "Melhor para cliques de feed"),
        @("Criativo B - Video Gancho 1",   "11/06/2026", "Ativo",    480.00, 48000, 1100, 95, "Gancho muito forte nos 3s"),
        @("Criativo C - Video Unboxing",   "12/06/2026", "Ativo",    160.00, 15000, 280,  28, "Excelente engajamento e retencao"),
        @("Criativo D - Imagem Depoimento","13/06/2026", "Pausado",  80.00,  8000,  120,  12, "Parado por CPA estourado no dia"),
        @("Criativo E - Video Promocional","14/06/2026", "Ativo",    240.00, 22000, 480,  40, "Visual moderno e dinamico")
    )
    
    for ($row = 5; $row -le 100; $row++) {
        $isSample = ($row -le 9)
        if ($isSample) {
            $data = $criativosData[$row - 5]
            Set-Cell $sheet3 $row 1  $data[0] -Align "left" -Bold $true
            Set-Cell $sheet3 $row 2  $data[1] -Format $fmtDate -Align "center"
            Set-Cell $sheet3 $row 3  $data[2] -Align "center"
            Set-Cell $sheet3 $row 4  $data[3] -Format $fmtCurrency
            Set-Cell $sheet3 $row 5  $data[4] -Format $fmtInteger
            Set-Cell $sheet3 $row 6  $data[5] -Format $fmtInteger
            Set-Cell $sheet3 $row 7  $data[6] -Format $fmtInteger
            Set-Cell $sheet3 $row 12 $data[7] -Align "left"
        } else {
            Set-Cell $sheet3 $row 1  "" -Align "left"
            Set-Cell $sheet3 $row 2  "" -Format $fmtDate -Align "center"
            Set-Cell $sheet3 $row 3  "" -Align "center"
            Set-Cell $sheet3 $row 4  "" -Format $fmtCurrency
            Set-Cell $sheet3 $row 5  "" -Format $fmtInteger
            Set-Cell $sheet3 $row 6  "" -Format $fmtInteger
            Set-Cell $sheet3 $row 7  "" -Format $fmtInteger
            Set-Cell $sheet3 $row 12 "" -Align "left"
        }
        
        Set-Cell $sheet3 $row 8  ('=IF(A' + $row + '="","",COUNTIFS(''Registro de Vendas''!D$5:D$1000, A' + $row + '))') -Format $fmtInteger
        Set-Cell $sheet3 $row 9  ('=IF(A' + $row + '="","",IF(H' + $row + '>0, D' + $row + '/H' + $row + ', 0))') -Format $fmtCurrency
        Set-Cell $sheet3 $row 10 ('=IF(A' + $row + '="","",IF(D' + $row + '>0, SUMIFS(''Registro de Vendas''!B$5:B$1000, ''Registro de Vendas''!D$5:D$1000, A' + $row + ')/D' + $row + ', 0))') -Format $fmtDecimal
        Set-Cell $sheet3 $row 11 ('=IF(A' + $row + '="","",IF(D' + $row + '=0, "Sem dados", IF(J' + $row + '>=2, "Excelente", IF(J' + $row + '>=1.5, "Bom", IF(J' + $row + '>=1, "Regular", "Ruim")))))') -Align "center" -Bold $true
        Set-Cell $sheet3 $row 13 ('=IF(A' + $row + '="", 0, J' + $row + ' + (1000-ROW())/1000000)')
        
        if ($row % 2 -eq 0) {
            $sheet3.Range("A$row:L$row").Interior.Color = $colorZebraBg
        }
        Set-Borders $sheet3 "A$row:L$row"
    }
    
    $sheet3.Columns.Item(13).Font.Color = Get-ExcelColor "#FFFFFF"
    
    $valStatusRange = $sheet3.Range("C5:C100")
    $valStatusRange.Validation.Delete()
    $valStatusRange.Validation.Add(3, 1, 1, ("Ativo" + $sep + "Pausado" + $sep + "Pendente" + $sep + "Descartado"))
    $valStatusRange.Validation.IgnoreBlank = $true
    $valStatusRange.Validation.InCellDropdown = $true
    
    $valCriativoVendas = $sheet6.Range("D5:D1000")
    $valCriativoVendas.Validation.Delete()
    $valCriativoVendas.Validation.Add(3, 1, 1, '=''Controle de Criativos''!$A$5:$A$100')
    $valCriativoVendas.Validation.IgnoreBlank = $true
    $valCriativoVendas.Validation.InCellDropdown = $true

    Set-Cell $sheet3 4 15 "Ranking de Criativos" -Bold $true -Size 11
    $sheet3.Range("O4:R4").Merge() | Out-Null
    Format-TableHeader $sheet3 "O4:R4"
    
    $rankingHeaders = @("Posicao", "Criativo", "ROAS", "CPA")
    for ($i = 0; $i -lt $rankingHeaders.Count; $i++) {
        Set-Cell $sheet3 5 ($i+15) $rankingHeaders[$i] -Bold $true -Align "center" -Color $colorCardBg
    }
    Set-Borders $sheet3 "O5:R5"
    
    Set-Cell $sheet3 6 15 '=UNICHAR(127942) & " Melhor"' -Bold $true -Align "left"
    Set-Cell $sheet3 6 16 '=IFERROR(INDEX(A$5:A$100, MATCH(LARGE(M$5:M$100, 1), M$5:M$100, 0)), "-")' -Align "left" -Bold $true
    Set-Cell $sheet3 6 17 '=IFERROR(INDEX(J$5:J$100, MATCH(LARGE(M$5:M$100, 1), M$5:M$100, 0)), 0)' -Format $fmtDecimal -Align "center"
    Set-Cell $sheet3 6 18 '=IFERROR(INDEX(I$5:I$100, MATCH(LARGE(M$5:M$100, 1), M$5:M$100, 0)), 0)' -Format $fmtCurrency -Align "right"
    Set-Borders $sheet3 "O6:R6"

    Set-Cell $sheet3 7 15 '=UNICHAR(129352) & " Segundo"' -Bold $true -Align "left"
    Set-Cell $sheet3 7 16 '=IFERROR(INDEX(A$5:A$100, MATCH(LARGE(M$5:M$100, 2), M$5:M$100, 0)), "-")' -Align "left" -Bold $true
    Set-Cell $sheet3 7 17 '=IFERROR(INDEX(J$5:J$100, MATCH(LARGE(M$5:M$100, 2), M$5:M$100, 0)), 0)' -Format $fmtDecimal -Align "center"
    Set-Cell $sheet3 7 18 '=IFERROR(INDEX(I$5:I$100, MATCH(LARGE(M$5:M$100, 2), M$5:M$100, 0)), 0)' -Format $fmtCurrency -Align "right"
    Set-Borders $sheet3 "O7:R7"

    Set-Cell $sheet3 8 15 '=UNICHAR(129353) & " Terceiro"' -Bold $true -Align "left"
    Set-Cell $sheet3 8 16 '=IFERROR(INDEX(A$5:A$100, MATCH(LARGE(M$5:M$100, 3), M$5:M$100, 0)), "-")' -Align "left" -Bold $true
    Set-Cell $sheet3 8 17 '=IFERROR(INDEX(J$5:J$100, MATCH(LARGE(M$5:M$100, 3), M$5:M$100, 0)), 0)' -Format $fmtDecimal -Align "center"
    Set-Cell $sheet3 8 18 '=IFERROR(INDEX(I$5:I$100, MATCH(LARGE(M$5:M$100, 3), M$5:M$100, 0)), 0)' -Format $fmtCurrency -Align "right"
    Set-Borders $sheet3 "O8:R8"
    
    $classRange = $sheet3.Range("K5:K100")
    $classRange.FormatConditions.Delete()
    $cfClassExc = $classRange.FormatConditions.Add(1, 3, '="Excelente"')
    $cfClassExc.Interior.Color = $colorSuccessBg
    $cfClassExc.Font.Color = $colorSuccessText
    $cfClassBom = $classRange.FormatConditions.Add(1, 3, '="Bom"')
    $cfClassBom.Interior.Color = $colorSuccessBg
    $cfClassBom.Font.Color = $colorSuccessText
    $cfClassReg = $classRange.FormatConditions.Add(1, 3, '="Regular"')
    $cfClassReg.Interior.Color = $colorWarningBg
    $cfClassReg.Font.Color = $colorWarningText
    $cfClassRuim = $classRange.FormatConditions.Add(1, 3, '="Ruim"')
    $cfClassRuim.Interior.Color = $colorCriticalBg
    $cfClassRuim.Font.Color = $colorCriticalText

    $sheet3.Activate()
    $sheet3.Range("A4:L4").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # =========================================================================
    # ABA 1: DASHBOARD EXECUTIVO
    # =========================================================================
    Write-Host "Configurando Dashboard Executivo..."
    Format-TitleHeader $sheet1 "B2:K2" "PAINEL EXECUTIVO - GESTAO DE TRAFEGO PAGO"
    
    function Style-KpiCard {
        param ($Sheet, $LabelCell, $ValueCell, $LabelText, $FormulaText, $FormatStr)
        Set-Cell $Sheet $LabelCell.Row $LabelCell.Column $LabelText -Bold $true -Align "center" -Color $colorCardBg -Size 9 -FontColor (Get-ExcelColor "#4B5563")
        Set-Cell $Sheet $ValueCell.Row $ValueCell.Column $FormulaText -Bold $true -Align "center" -Format $FormatStr -Size 13
        $Sheet.Range($Sheet.Cells.Item($LabelCell.Row, $LabelCell.Column), $Sheet.Cells.Item($LabelCell.Row, $LabelCell.Column+1)).Merge() | Out-Null
        $Sheet.Range($Sheet.Cells.Item($ValueCell.Row, $ValueCell.Column), $Sheet.Cells.Item($ValueCell.Row, $ValueCell.Column+1)).Merge() | Out-Null
        $boxAddr = $Sheet.Cells.Item($LabelCell.Row, $LabelCell.Column).Address() + ":" + $Sheet.Cells.Item($ValueCell.Row, $ValueCell.Column+1).Address()
        Set-Borders $Sheet $boxAddr
    }
    
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(4, 2)) ($sheet1.Cells.Item(5, 2)) "Investimento Total" "=SUM('Controle Diario'!B$5:B$100)" $fmtCurrency
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(4, 4)) ($sheet1.Cells.Item(5, 4)) "Faturamento Total" "=SUM('Controle Diario'!I$5:I$100)" $fmtCurrency
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(4, 6)) ($sheet1.Cells.Item(5, 6)) "Lucro Bruto" "=D5-B5" $fmtCurrency
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(4, 8)) ($sheet1.Cells.Item(5, 8)) "ROAS Medio" "=IF(B5>0, D5/B5, 0)" $fmtDecimal
    
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(7, 2)) ($sheet1.Cells.Item(8, 2)) "Compras Totais" "=SUM('Controle Diario'!H$5:H$100)" $fmtInteger
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(7, 4)) ($sheet1.Cells.Item(8, 4)) "CPA Medio" "=IF(B8>0, B5/B8, 0)" $fmtCurrency
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(7, 6)) ($sheet1.Cells.Item(8, 6)) "Ticket Medio" "=IF(B8>0, D5/B8, 0)" $fmtCurrency
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(7, 8)) ($sheet1.Cells.Item(8, 8)) "Cliques Totais" "=SUM('Controle Diario'!E$5:E$100)" $fmtInteger
    
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(10, 2)) ($sheet1.Cells.Item(11, 2)) "Finalizacoes de Compra" "=SUM('Controle Diario'!G$5:G$100)" $fmtInteger
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(10, 4)) ($sheet1.Cells.Item(11, 4)) "Tx. Landing -> Checkout" "=IF(SUM('Controle Diario'!F$5:F$100)>0, B11/SUM('Controle Diario'!F$5:F$100), 0)" $fmtPercent
    Style-KpiCard $sheet1 ($sheet1.Cells.Item(10, 6)) ($sheet1.Cells.Item(11, 6)) "Tx. Checkout -> Compra" "=IF(B11>0, B8/B11, 0)" $fmtPercent
    
    Set-Cell $sheet1 4 10 "Metas Desejadas" -Bold $true -Align "center" -Color $colorHeaderBg -FontColor 16777215
    $sheet1.Range("J4:K4").Merge() | Out-Null
    
    Set-Cell $sheet1 5 10 "Meta ROAS" -Bold $true -Align "left" -Color $colorCardBg
    Set-Cell $sheet1 5 11 2.0 -Format $fmtDecimal -Bold $true -Align "center"
    
    Set-Cell $sheet1 6 10 "Meta CPA Max" -Bold $true -Align "left" -Color $colorCardBg
    Set-Cell $sheet1 6 11 20.00 -Format $fmtCurrency -Bold $true -Align "center"
    Set-Borders $sheet1 "J4:K6"
    
    Set-Cell $sheet1 8 10 "Status ROAS" -Bold $true -Align "left" -Color $colorCardBg
    Set-Cell $sheet1 8 11 '=IF(H5<1, UNICHAR(128992) & " Acao", IF(H5<K5, UNICHAR(128993) & " Atencao", UNICHAR(128994) & " OK"))' -Bold $true -Align "center"
    
    Set-Cell $sheet1 9 10 "Status CPA" -Bold $true -Align "left" -Color $colorCardBg
    Set-Cell $sheet1 9 11 '=IF(D8=0, "-", IF(D8>K6*1.5, UNICHAR(128992) & " Acao", IF(D8>K6, UNICHAR(128993) & " Atencao", UNICHAR(128994) & " OK")))' -Bold $true -Align "center"
    Set-Borders $sheet1 "J8:K9"

    # =========================================================================
    # ABA 4: ANALISE DE FUNIL
    # =========================================================================
    Write-Host "Configurando Analise de Funil..."
    Format-TitleHeader $sheet4 "B2:F2" "ANALISE DO FUNIL DE VENDAS"
    
    $colsFunil = @("Etapa do Funil", "Volume Total", "Taxa da Etapa", "Taxa Acumulada", "Alerta / Diagnostico Operacional")
    for ($i = 0; $i -lt $colsFunil.Count; $i++) {
        Set-Cell $sheet4 4 ($i+2) $colsFunil[$i]
    }
    Format-TableHeader $sheet4 "B4:F4"
    
    $etapas = @("1. Impressoes", "2. Cliques no Link", "3. Visualizacoes da Pagina", "4. Finalizacoes de Compra", "5. Compras Realizadas")
    for ($i = 0; $i -lt 5; $i++) {
        $r = $i + 5
        Set-Cell $sheet4 $r 2 $etapas[$i] -Bold $true -Align "left"
    }
    
    Set-Cell $sheet4 5 3 "=SUM('Controle Diario'!C$5:C$100)" -Format $fmtInteger -Align "right"
    Set-Cell $sheet4 6 3 "=SUM('Controle Diario'!E$5:E$100)" -Format $fmtInteger -Align "right"
    Set-Cell $sheet4 7 3 "=SUM('Controle Diario'!F$5:F$100)" -Format $fmtInteger -Align "right"
    Set-Cell $sheet4 8 3 "=SUM('Controle Diario'!G$5:G$100)" -Format $fmtInteger -Align "right"
    Set-Cell $sheet4 9 3 "=SUM('Controle Diario'!H$5:H$100)" -Format $fmtInteger -Align "right"
    
    Set-Cell $sheet4 5 4 1.0 -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 6 4 "=IF(C5>0, C6/C5, 0)" -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 7 4 "=IF(C6>0, C7/C6, 0)" -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 8 4 "=IF(C7>0, C8/C7, 0)" -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 9 4 "=IF(C8>0, C9/C8, 0)" -Format $fmtPercent -Align "center"
    
    Set-Cell $sheet4 5 5 1.0 -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 6 5 "=IF(C5>0, C6/C5, 0)" -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 7 5 "=IF(C5>0, C7/C5, 0)" -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 8 5 "=IF(C5>0, C8/C5, 0)" -Format $fmtPercent -Align "center"
    Set-Cell $sheet4 9 5 "=IF(C5>0, C9/C5, 0)" -Format $fmtPercent -Align "center"
    
    Set-Cell $sheet4 5 6 "Volume de entrada do funil" -Align "left"
    Set-Cell $sheet4 6 6 '=IF(D6<0.01, UNICHAR(128992) & " Criativo (CTR < 1% indica baixa atratividade)", UNICHAR(128994) & " CTR saudavel nos criativos")' -Align "left"
    Set-Cell $sheet4 7 6 '=IF(D7<0.7, UNICHAR(128992) & " Carregamento (Queda > 30% - verifique velocidade)", UNICHAR(128994) & " Carregamento saudavel")' -Align "left"
    Set-Cell $sheet4 8 6 '=IF(D8<0.05, UNICHAR(128992) & " Landing Page (conversao para checkout < 5%)", UNICHAR(128994) & " Landing Page saudavel")' -Align "left"
    Set-Cell $sheet4 9 6 '=IF(D9<0.05, UNICHAR(128992) & " Checkout/Oferta (conversao checkout < 5%)", UNICHAR(128994) & " Checkout/Oferta saudavel")' -Align "left"
    
    for ($r = 5; $r -le 9; $r++) {
        Set-Borders $sheet4 "B$r:F$r"
        if ($r % 2 -eq 0) {
            $sheet4.Range("B$r:F$r").Interior.Color = $colorZebraBg
        }
    }
    
    $funilAlertRange = $sheet4.Range("F6:F9")
    $funilAlertRange.FormatConditions.Delete()
    
    $cfFunilRed = $funilAlertRange.FormatConditions.Add(2, $null, '=ESQUERDA(F6;1)=CARACT.UNICODE(128992)')
    $cfFunilRed.Interior.Color = $colorCriticalBg
    $cfFunilRed.Font.Color = $colorCriticalText
    
    $cfFunilGreen = $funilAlertRange.FormatConditions.Add(2, $null, '=ESQUERDA(F6;1)=CARACT.UNICODE(128994)')
    $cfFunilGreen.Interior.Color = $colorSuccessBg
    $cfFunilGreen.Font.Color = $colorSuccessText

    # =========================================================================
    # ABA 5: HISTORICO DE ESCALA
    # =========================================================================
    Write-Host "Configurando Historico de Escala..."
    Format-TitleHeader $sheet5 "A2:H2" "HISTORICO E REGISTRO DE ESCALA"
    
    $colsEscala = @("Data", "Campanha", "Orcamento Anterior", "Novo Orcamento", "Percentual de Escala", 
                    "Motivo da Alteracao", "Resultado Apos Escala", "Observacoes")
    for ($i = 0; $i -lt $colsEscala.Count; $i++) {
        Set-Cell $sheet5 4 ($i+1) $colsEscala[$i]
    }
    Format-TableHeader $sheet5 "A4:H4"
    
    $escalaData = @(
        @("12/06/2026", "Campanha Validacao", 50.00,  75.00,  "ROI Alto", "Positivo (Manteve ROI)", "CPA estavel sob R$ 15"),
        @("15/06/2026", "Campanha Validacao", 75.00,  100.00, "ROI Alto", "Positivo (Manteve ROI)", "Escala vertical bem sucedida"),
        @("18/06/2026", "Campanha Validacao", 100.00, 120.00, "CPA Abaixo da Meta", "Neutro (CPA aceitavel)", "Subiu orcamento de seguranca")
    )
    
    for ($row = 5; $row -le 100; $row++) {
        $isSample = ($row -le 7)
        if ($isSample) {
            $data = $escalaData[$row - 5]
            Set-Cell $sheet5 $row 1 $data[0] -Format $fmtDate -Align "center"
            Set-Cell $sheet5 $row 2 $data[1] -Align "left" -Bold $true
            Set-Cell $sheet5 $row 3 $data[2] -Format $fmtCurrency
            Set-Cell $sheet5 $row 4 $data[3] -Format $fmtCurrency
            Set-Cell $sheet5 $row 6 $data[4] -Align "left"
            Set-Cell $sheet5 $row 7 $data[5] -Align "center"
            Set-Cell $sheet5 $row 8 $data[6] -Align "left"
        } else {
            Set-Cell $sheet5 $row 1 "" -Format $fmtDate -Align "center"
            Set-Cell $sheet5 $row 2 "" -Align "left"
            Set-Cell $sheet5 $row 3 "" -Format $fmtCurrency
            Set-Cell $sheet5 $row 4 "" -Format $fmtCurrency
            Set-Cell $sheet5 $row 6 "" -Align "left"
            Set-Cell $sheet5 $row 7 "" -Align "center"
            Set-Cell $sheet5 $row 8 "" -Align "left"
        }
        
        Set-Cell $sheet5 $row 5 ('=IF(A' + $row + '="","",IF(C' + $row + '>0, (D' + $row + '-C' + $row + ')/C' + $row + ', 0))') -Format $fmtPercent -Align "center"
        
        if ($row % 2 -eq 0) {
            $sheet5.Range("A$row:H$row").Interior.Color = $colorZebraBg
        }
        Set-Borders $sheet5 "A$row:H$row"
    }
    
    $valMotivos = $sheet5.Range("F5:F100")
    $valMotivos.Validation.Delete()
    $valMotivos.Validation.Add(3, 1, 1, ("ROI Alto" + $sep + "CPA Abaixo da Meta" + $sep + "Validacao de Escala" + $sep + "Reducao de Custos"))
    $valMotivos.Validation.IgnoreBlank = $true
    $valMotivos.Validation.InCellDropdown = $true
    
    $valResultados = $sheet5.Range("G5:G100")
    $valResultados.Validation.Delete()
    $valResultados.Validation.Add(3, 1, 1, ("Positivo (Manteve ROI)" + $sep + "Neutro (CPA aceitavel)" + $sep + "Negativo (CPA estourou)" + $sep + "Em Analise"))
    $valResultados.Validation.IgnoreBlank = $true
    $valResultados.Validation.InCellDropdown = $true
    
    $sheet5.Activate()
    $sheet5.Range("A4:H4").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # =========================================================================
    # ABA 7: INSIGHTS AUTOMATICOS
    # =========================================================================
    Write-Host "Configurando Insights Automaticos..."
    Format-TitleHeader $sheet7 "B2:D2" "INSIGHTS OPERACIONAIS AUTOMATICOS"
    
    $colsInsights = @("Area de Analise", "Diagnostico", "Recomendacao Operacional")
    for ($i = 0; $i -lt $colsInsights.Count; $i++) {
        Set-Cell $sheet7 4 ($i+2) $colsInsights[$i]
    }
    Format-TableHeader $sheet7 "B4:D4"
    
    Set-Cell $sheet7 5 2 "CPA vs Meta" -Bold $true
    Set-Cell $sheet7 5 3 '=IF(''Dashboard Executivo''!D8=0, "Sem dados de vendas.", IF(''Dashboard Executivo''!D8<=''Dashboard Executivo''!K6, UNICHAR(128994) & " CPA dentro da meta.", UNICHAR(128992) & " CPA acima da meta."))' -Bold $true
    Set-Cell $sheet7 5 4 '=IF(''Dashboard Executivo''!D8=0, "-", IF(''Dashboard Executivo''!D8<=''Dashboard Executivo''!K6, "Sua aquisicao esta saudavel. Pode manter ou escalar.", "Cuidado! Pause criativos ruins e reduza verba para controlar custos."))'
    
    Set-Cell $sheet7 6 2 "ROAS vs Meta" -Bold $true
    Set-Cell $sheet7 6 3 '=IF(''Dashboard Executivo''!H5=0, "Sem dados de investimento.", IF(''Dashboard Executivo''!H5>=''Dashboard Executivo''!K5, UNICHAR(128994) & " ROAS dentro da meta.", UNICHAR(128992) & " ROAS abaixo da meta."))' -Bold $true
    Set-Cell $sheet7 6 4 '=IF(''Dashboard Executivo''!H5=0, "-", IF(''Dashboard Executivo''!H5>=''Dashboard Executivo''!K5, "Operacao lucrativa. Bom momento para escala gradual.", "Trabalhando no vermelho. Revise o preco ou teste novos criativos."))'
    
    Set-Cell $sheet7 7 2 "Conversao de Checkout" -Bold $true
    Set-Cell $sheet7 7 3 '=IF(''Analise de Funil''!D9<0.05, UNICHAR(128992) & " Checkout baixo (<5%)", UNICHAR(128994) & " Checkout saudavel")' -Bold $true
    Set-Cell $sheet7 7 4 '=IF(''Analise de Funil''!D9<0.05, "Gargalo no checkout. Use checkout de 1 clique, Pix automatico e recupere abandonos.", "Mantenha a estrutura e continue otimizando.")'
    
    Set-Cell $sheet7 8 2 "Conversao de Landing Page" -Bold $true
    Set-Cell $sheet7 8 3 '=IF(''Analise de Funil''!D8<0.05, UNICHAR(128992) & " Landing Page baixa (<5%)", UNICHAR(128994) & " Landing Page saudavel")' -Bold $true
    Set-Cell $sheet7 8 4 '=IF(''Analise de Funil''!D8<0.05, "Verifique velocidade e garanta promessa clara na primeira dobra.", "Mantenha e faca testes A/B pontuais.")'
    
    Set-Cell $sheet7 9 2 "Atratividade dos Criativos" -Bold $true
    Set-Cell $sheet7 9 3 '=IF(''Analise de Funil''!D6<0.01, UNICHAR(128992) & " CTR baixo (<1%)", UNICHAR(128994) & " CTR saudavel")' -Bold $true
    Set-Cell $sheet7 9 4 '=IF(''Analise de Funil''!D6<0.01, "Produza criativos com hooks (ganchos) de 3 segundos mais fortes.", "Parabens! Continue produzindo variacoes dos melhores.")'
    
    Set-Cell $sheet7 10 2 "Tendencia de CPA" -Bold $true
    Set-Cell $sheet7 10 3 '=IF(COUNTA(''Controle Diario''!A$5:A$100)<7, "Historico insuficiente.", IF(AND(INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)) > INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)-1), INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)-1) > INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)-2)), UNICHAR(128992) & " CPA subindo ha 3 dias", UNICHAR(128994) & " CPA saudavel/estavel"))' -Bold $true
    Set-Cell $sheet7 10 4 '=IF(COUNTA(''Controle Diario''!A$5:A$100)<7, "-", IF(AND(INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)) > INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)-1), INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)-1) > INDEX(''Controle Diario''!L$5:L$100, COUNTA(''Controle Diario''!A$5:A$100)-2)), "Fadiga de criativo provavel. Renove os anuncios do conjunto.", "Normalidade operacional mantida."))'
    
    Set-Cell $sheet7 11 2 "Criativo Destaque" -Bold $true
    Set-Cell $sheet7 11 3 '=UNICHAR(127942) & " Criativo: " & ''Controle de Criativos''!P6' -Bold $true
    Set-Cell $sheet7 11 4 '=CONCATENATE("ROAS de ", TEXT(''Controle de Criativos''!Q6, "0.0"), ". Considere criar variacoes do mesmo.")'
    
    for ($r = 5; $r -le 11; $r++) {
        Set-Borders $sheet7 "B$r:D$r"
        if ($r % 2 -eq 0) {
            $sheet7.Range("B$r:D$r").Interior.Color = $colorZebraBg
        }
    }
    
    $insightsRange = $sheet7.Range("C5:C11")
    $insightsRange.FormatConditions.Delete()
    
    $cfInsRed = $insightsRange.FormatConditions.Add(2, $null, '=ESQUERDA(C5;1)=CARACT.UNICODE(128992)')
    $cfInsRed.Interior.Color = $colorCriticalBg
    $cfInsRed.Font.Color = $colorCriticalText
    
    $cfInsYellow = $insightsRange.FormatConditions.Add(2, $null, '=ESQUERDA(C5;1)=CARACT.UNICODE(128993)')
    $cfInsYellow.Interior.Color = $colorWarningBg
    $cfInsYellow.Font.Color = $colorWarningText
    
    $cfInsGreen = $insightsRange.FormatConditions.Add(2, $null, '=OU(ESQUERDA(C5;1)=CARACT.UNICODE(128994);ESQUERDA(C5;1)=CARACT.UNICODE(127942))')
    $cfInsGreen.Interior.Color = $colorSuccessBg
    $cfInsGreen.Font.Color = $colorSuccessText

    # =========================================================================
    # CRIANDO OS GRAFICOS
    # =========================================================================
    Write-Host "Criando Graficos..."
    $xlLine = 4
    $xlColumnClustered = 51
    
    $sheet1.Activate()
    $chart1Obj = $sheet1.ChartObjects().Add(20, 230, 420, 240)
    $chart1 = $chart1Obj.Chart
    $chart1.ChartType = $xlLine
    $chart1Range = $sheet2.Range("A4:B14;I4:I14")
    $chart1.SetSourceData($chart1Range)
    $chart1.HasTitle = $true
    $chart1.ChartTitle.Text = "Evolucao Diaria de Gastos vs. Faturamento"
    $chart1.ChartTitle.Font.Name = "Segoe UI"
    $chart1.ChartTitle.Font.Size = 10
    $chart1.ChartTitle.Font.Bold = $true
    
    $chart2Obj = $sheet1.ChartObjects().Add(460, 230, 420, 240)
    $chart2 = $chart2Obj.Chart
    $chart2.ChartType = $xlColumnClustered
    $chart2Range = $sheet2.Range("A4:A14;H4:H14")
    $chart2.SetSourceData($chart2Range)
    $chart2.HasTitle = $true
    $chart2.ChartTitle.Text = "Evolucao Diaria de Compras"
    $chart2.ChartTitle.Font.Name = "Segoe UI"
    $chart2.ChartTitle.Font.Size = 10
    $chart2.ChartTitle.Font.Bold = $true
    
    $chart3Obj = $sheet1.ChartObjects().Add(20, 480, 860, 240)
    $chart3 = $chart3Obj.Chart
    $chart3.ChartType = $xlLine
    $chart3Range = $sheet2.Range("A4:A14;L4:M14")
    $chart3.SetSourceData($chart3Range)
    $chart3.HasTitle = $true
    $chart3.ChartTitle.Text = "Tendencia de CPA vs. ROAS Diario"
    $chart3.ChartTitle.Font.Name = "Segoe UI"
    $chart3.ChartTitle.Font.Size = 10
    $chart3.ChartTitle.Font.Bold = $true
    
    try {
        $roasSeries = $chart3.SeriesCollection(2)
        $roasSeries.AxisGroup = 2
    } catch {
        Write-Host "Aviso: Nao foi possivel colocar o ROAS no eixo secundario. Continuando..."
    }
 
    $sheet5.Activate()
    $chartScaleObj = $sheet5.ChartObjects().Add(20, 180, 750, 240)
    $chartScale = $chartScaleObj.Chart
    $chartScale.ChartType = $xlColumnClustered
    $chartScaleRange = $sheet5.Range("A4:B7;D4:D7")
    $chartScale.SetSourceData($chartScaleRange)
    $chartScale.HasTitle = $true
    $chartScale.ChartTitle.Text = "Evolucao do Orcamento de Escala"
    $chartScale.ChartTitle.Font.Name = "Segoe UI"
    $chartScale.ChartTitle.Font.Size = 10
    $chartScale.ChartTitle.Font.Bold = $true

    # =========================================================================
    # POLIMENTO DE CADA ABA
    # =========================================================================
    Write-Host "Polindo largura das colunas..."
    foreach ($sh in @($sheet1, $sheet2, $sheet3, $sheet4, $sheet5, $sheet6, $sheet7)) {
        $sh.UsedRange.Columns.AutoFit() | Out-Null
        
        if ($sh.Name -eq "Controle Diario") {
            $sh.Columns.Item(16).ColumnWidth = 28
        }
        if ($sh.Name -eq "Controle de Criativos") {
            $sh.Columns.Item(1).ColumnWidth = 25
            $sh.Columns.Item(12).ColumnWidth = 28
        }
        if ($sh.Name -eq "Registro de Vendas") {
            $sh.Columns.Item(4).ColumnWidth = 25
            $sh.Columns.Item(6).ColumnWidth = 25
        }
        if ($sh.Name -eq "Insights Automaticos") {
            $sh.Columns.Item(2).ColumnWidth = 22
            $sh.Columns.Item(3).ColumnWidth = 32
            $sh.Columns.Item(4).ColumnWidth = 55
        }
        if ($sh.Name -eq "Analise de Funil") {
            $sh.Columns.Item(6).ColumnWidth = 55
        }
    }
    
    $workspacePath = Join-Path $PSScriptRoot "Gestao_Trafego_Low_Ticket.xlsx"
    Write-Host "Salvando planilha em: $workspacePath"
    $workbook.SaveAs($workspacePath) | Out-Null
    
    Write-Host "Planilha gerada com sucesso!"
    
} catch {
    Write-Host "Erro ocorrido: $_"
    throw $_
} finally {
    if ($workbook) {
        $workbook.Close($false) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
