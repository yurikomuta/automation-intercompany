# Script para executar testes com UTF-8 adequado
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Remover arquivo antigo se existir
if (Test-Path 'test-results.txt') {
    Remove-Item 'test-results.txt'
}

# Executar testes e capturar saída
& {
    . './tests/SubjectParser.Tests.ps1'
} *> 'test-results.txt'

# Exibir resultados com encoding UTF-8
Get-Content -Path 'test-results.txt' -Encoding UTF8
