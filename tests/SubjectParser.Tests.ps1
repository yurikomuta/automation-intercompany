Import-Module `
    ".\src\SubjectParser.psm1" `
    -Force

$subject = @(
    'Relatório de Eliminação Patrimonial',
    'TBRA (0185) X ITIS (0306)',
    '08/2026'
)-join ' - '

$result = ConvertFrom-PatrimonialSubject -Subject $subject

$result | Format-List
