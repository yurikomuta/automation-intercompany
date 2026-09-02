Set-StrictMode -Version Latest

# PadrÃ£o oficial do assunto:
# RelatÃ³rio de EliminaÃ§Ã£o Patrimonial -
# SIGLA (CÃ“DIGO) X SIGLA (CÃ“DIGO) - MM/AAAA

$script:PatrimonialSubjectPattern = @(
    '^RelatÃ³rio de EliminaÃ§Ã£o Patrimonial'
    '\s*-\s*'
    '(?<Company1Acronym>[A-Z0-9]{2,15})'
    '\s*\(\s*'
    '(?<Company1Code>\d{4})'
    '\s*\)'
    '\s*X\s*'
    '(?<Company2Acronym>[A-Z0-9]{2,15})'
    '\s*\(\s*'
    '(?<Company2Code>\d{4})'
    '\s*\)'
    '\s*-\s*'
    '(?<Month>0[1-9]|1[0-2])'
    '/'
    '(?<Year>\d{4})'
    '$'
) -join ''

function ConvertFrom-PatrimonialSubject {
    <#
    .SYNOPSIS
        Interpreta o assunto de um e-mail patrimonial.

    .DESCRIPTION
        Valida o assunto, extrai os cÃ³digos, as siglas e a competÃªncia
        e gera a chave utilizada para consultar o routes.json.

    .PARAMETER Subject
        Assunto original recebido do Outlook.

    .PARAMETER AcceptReplyPrefixes
        Permite remover prefixos como RE:, RES:, ENC:, FW: e FWD:
        antes de validar o assunto.

    .OUTPUTS
        PSCustomObject do tipo Patrimonial.SubjectInfo.

    .EXAMPLE
        ConvertFrom-PatrimonialSubject `
            -Subject "RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - TBRA (0185) X ITIS (0306) - 08/2026"
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [Parameter()]
        [switch]$AcceptReplyPrefixes
    )

    process {
        # Substitui espaÃ§o nÃ£o separÃ¡vel, que pode aparecer em textos
        # copiados de HTML ou gerados automaticamente.
        $normalizedSubject = $Subject.Replace(
            [char]0x00A0,
            [char]0x0020
        )

        # Normaliza caracteres Unicode, incluindo os acentos.
        $normalizedSubject = $normalizedSubject.Normalize(
            [System.Text.NormalizationForm]::FormC
        )

        # Remove espaÃ§os nas extremidades e transforma vÃ¡rios espaÃ§os em um.
        $normalizedSubject = $normalizedSubject.Trim()
        $normalizedSubject = [regex]::Replace(
            $normalizedSubject,
            '\s+',
            ' '
        )

        if ([string]::IsNullOrWhiteSpace($normalizedSubject)) {
            $exception = [System.FormatException]::new(
                'O assunto estÃ¡ vazio apÃ³s sua normalizaÃ§Ã£o.'
            )

            $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                $exception,
                'EMPTY_SUBJECT',
                [System.Management.Automation.ErrorCategory]::InvalidData,
                $Subject
            )

            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }

        $replyPrefixRemoved = $false
        $subjectForParsing = $normalizedSubject

        if ($AcceptReplyPrefixes.IsPresent) {
            $prefixPattern = '^(?:(?:RE|RES|ENC|FW|FWD)\s*:\s*)+'

            $subjectWithoutPrefix = [regex]::Replace(
                $subjectForParsing,
                $prefixPattern,
                '',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            ).Trim()

            $replyPrefixRemoved = (
                $subjectWithoutPrefix -ne $subjectForParsing
            )

            $subjectForParsing = $subjectWithoutPrefix
        }

        $regexOptions = (
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase `
            -bor
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
        )

        $subjectMatch = [regex]::Match(
            $subjectForParsing,
            $script:PatrimonialSubjectPattern,
            $regexOptions
        )

        if (-not $subjectMatch.Success) {
            $expectedFormat = (
                'RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'SIGLA (0000) X SIGLA (0000) - MM/AAAA'
            )

            $message = (
                "O assunto nao corresponde ao padrao esperado. " +
                "Assunto recebido: '$Subject'. " +
                "Formato esperado: '$expectedFormat'."
            )

            $exception = [System.FormatException]::new($message)

            $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                $exception,
                'INVALID_SUBJECT_FORMAT',
                [System.Management.Automation.ErrorCategory]::InvalidData,
                $Subject
            )

            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }

        # CÃ³digos permanecem como string para preservar zeros Ã  esquerda.
        $company1Code = $subjectMatch.Groups[
            'Company1Code'
        ].Value

        $company2Code = $subjectMatch.Groups[
            'Company2Code'
        ].Value

        $company1Acronym = $subjectMatch.Groups[
            'Company1Acronym'
        ].Value.ToUpperInvariant()

        $company2Acronym = $subjectMatch.Groups[
            'Company2Acronym'
        ].Value.ToUpperInvariant()

        $month = [int]$subjectMatch.Groups['Month'].Value
        $year = [int]$subjectMatch.Groups['Year'].Value

        $competence = '{0:D2}/{1:D4}' -f $month, $year

        $routeKey = '{0}:{1}|{2}:{3}' -f `
            $company1Code,
            $company1Acronym,
            $company2Code,
            $company2Acronym

        [PSCustomObject]@{
            PSTypeName        = 'Patrimonial.SubjectInfo'
            OriginalSubject   = $Subject
            NormalizedSubject = $subjectForParsing
            Company1Code      = $company1Code
            Company1Acronym   = $company1Acronym
            Company2Code      = $company2Code
            Company2Acronym   = $company2Acronym
            Month             = $month
            Year              = $year
            Competence        = $competence
            RouteKey          = $routeKey
            ReplyPrefixRemoved = $replyPrefixRemoved
        }
    }
}

Export-ModuleMember -Function ConvertFrom-PatrimonialSubject
