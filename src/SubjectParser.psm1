
Set-StrictMode -Version Latest

#Padrão oficial do assunto: {Company1Code}:{Company1Acronym}|{Company2Code}:{Company2Acronym} - {Competence}

# Relatório de Eliminação Patromonial - SIGLA (CÓDIGO) X SIGLA (CÓDIGO) - MM/YYYY

$script:PatrimonialSubjectPattern = @(
    '^Relatório de Eliminação Patrimonial'
    '\s*-\s*'
    '(?<Company1Acronym>[A-Z0-9]{2,15})'
    '\s*\(\s*'
    '(?<Company1Code>\d{4})'
    '\s*)'
    '\s*X\s*'
    '(?<Company2Acronym>[A-Z0-9]{2,15})'
    '\s*\(\s*'
    '(?<Company2Code>\d{4})'
    '\s*\)\s*'
    '\s*-\s*'
    '(?<Month>0[1-9]|1[0-2])'
    '/'
    '(?<Year>\d{4})'
    '$'
) -join ''

function ConvertFrom-PatrimonialSubject{
    <#
    .SYNOPSIS
        Interpretar o assunto de um e-mail patrimonial
    .DESCRIPTION
    Valida o assunto, extrai  os códigos, as siglas e a competência e gera a chave para consultar o route.json
    .PARAMETER Subject
    O assunto original recebido no Outlook.
    .PARAMETER AcceptReplyPrefixes
    Aceita prefixos de resposta, como "RE:" e "FW:", para que o assunto seja interpretado mesmo quando o e-mail for uma resposta ou encaminhamento.
    .OUTPUTS
        PSCustomObject do tipo Patrimonial.SubjectInfo.
    .EXAMPLE
    ConvertFrom-PatrimonialSubject 
        -Subject "Relatório de Eliminação Patrimonial - TBRA (0185) X ITIS (0306) - 08/2026"
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(
            Mandatory = $true,
            Position= 0,
            ValueFromPipeline = $true
        )]

        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [Parameter()]
        [switch]$AccepReplyPrefixes        
    )

    process{
        #Substitui os espaços, 
        $normalizedSubject = $Subject.Replace(
            [char]0x00A0,
            [char]0x0020
        )

        #Normaliza os caracteres Unicode
        $normalizedSubject = $normalizedSubject.Normalize(
            [System.Text.NormalizationForm]::FormC
        )
        
        #Remove espaços nas extremidades e transforma vários espaços em um único espaço
        $normalizedSubject = $normalizedSubject.Trim()
        $normalizedSubject = [regex]::Replace(
            $normalizedSubject,
            '\s+',
            ' '
        )

        if ([string]::IsNullOrWhiteSpace($normalizedSubject)) {
            $exception = [System.FormatException]::new(
                "O assunto está vazio após sua normalização."
                
            )

            $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                $exception,
                "EMPTY_SUBJECT",
                [System.Management.Automation.ErrorCategory]::InvalidData,
                $Subject
            )

            SPSCmlet.ThrowTerminatingError($errorRecord)
        }

        $replyPrefixRemoved = $false
        $subjectForParsing = $normalizedSubject

        if ($AccepReplyPrefixes.IsPresent){
            $prefixPattern = '^(?:(?:RE|RES|ENC|FW|FWD)\s*:\s*)+'

            $subjectForParsing = [regex]::Replace(
                $subjectForParsing,
                $prefixPattern,
                '',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            ).Trim()

            $replyPrefixRemoved = (
                $subjectWithoutPrefix -ne $subjectForParsing
            )
        }



    }

}



[PSCustomObject]@{
    Company1Code = "0185"
    Company1Acronym = "TBRA"
    Company2Code = "0306"
    Company2Acronym = "ITIS"
    Month = 8
    Year = 2026
    Competence = "08/2026"
    RouteKey = "0185:TBRA|0306:ITIS"
}