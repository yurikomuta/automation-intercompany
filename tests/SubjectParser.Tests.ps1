# Configurar codificaÃ§Ã£o UTF-8 
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Import-Module `
    "./src/SubjectParser.psm1" `
    -Force  

Describe 'ConvertFrom-PatrimonialSubject' {

    Context 'Quando o assunto Ã© vÃ¡lido' {

        It 'Deve extrair corretamente todas as informaÃ§Ãµes' {
            $subject = (
                'RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject

            $result.Company1Acronym | Should -Be 'TBRA'
            $result.Company1Code | Should -Be '0185'

            $result.Company2Acronym | Should -Be 'ITIS'
            $result.Company2Code | Should -Be '0306'

            $result.Month | Should -Be 8
            $result.Year | Should -Be 2026
            $result.Competence | Should -Be '08/2026'

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'
        }

        It 'Deve normalizar espaÃ§os adicionais' {
            $subject = (
                '  RelatÃ³rio de EliminaÃ§Ã£o Patrimonial  -  ' +
                'TBRA  (0185)  X  ITIS (0306)  -  08/2026  '
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'

            $result.Competence |
                Should -Be '08/2026'
        }

        It 'Deve converter as siglas para letras maiÃºsculas' {
            $subject = (
                'RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'tbra (0185) X itis (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject

            $result.Company1Acronym | Should -Be 'TBRA'
            $result.Company2Acronym | Should -Be 'ITIS'
        }
    }

    Context 'Quando o assunto Ã© invÃ¡lido' {

        It 'Deve lanÃ§ar erro quando o formato estÃ¡ incorreto' {
            $subject = 'RelatÃ³rio incorreto - TBRA X ITIS'

            {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject
            } | Should -Throw
        }

        It 'Deve retornar INVALID_SUBJECT_FORMAT para assunto incorreto' {
            $subject = 'RelatÃ³rio incorreto - TBRA X ITIS'

            $capturedError = $null

            try {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject `
                    -ErrorAction Stop |
                    Out-Null
            }
            catch {
                $capturedError = $_
            }

            # Confirma que um erro realmente foi capturado.
            $capturedError |
                Should -Not -BeNullOrEmpty

            # Confirma o identificador criado no SubjectParser.
            $capturedError.FullyQualifiedErrorId |
                Should -Match '^INVALID_SUBJECT_FORMAT'

            # Confirma a categoria do erro.
            $capturedError.CategoryInfo.Category |
                Should -Be 'InvalidData'

            # Confirma o tipo da exceção.
            $capturedError.Exception.GetType().FullName |
                Should -Be 'System.FormatException'

            # Confirma parte da mensagem.
            $capturedError.Exception.Message |
                Should -Match 'nao corresponde ao padrao esperado'
        }

        It 'Deve rejeitar mÃªs maior que doze' {
            $subject = (
                'RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 13/2026'
            )

            $capturedError = $null

            try {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject `
                    -ErrorAction Stop |
                    Out-Null
            }
            catch {
                $capturedError = $_
            }

            $capturedError |
                Should -Not -BeNullOrEmpty

            $capturedError.FullyQualifiedErrorId |
                Should -Match '^INVALID_SUBJECT_FORMAT'
        }

        It 'Deve rejeitar cÃ³digo com menos de quatro dÃ­gitos' {
            $subject = (
                'RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'TBRA (185) X ITIS (0306) - 08/2026'
            )

            {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject
            } | Should -Throw
        }

        It 'Deve retornar EMPTY_SUBJECT para assunto com espaÃ§os' {
            $capturedError = $null

            try {
                ConvertFrom-PatrimonialSubject `
                    -Subject '   ' `
                    -ErrorAction Stop |
                    Out-Null
            }
            catch {
                $capturedError = $_
            }

            $capturedError |
                Should -Not -BeNullOrEmpty

            $capturedError.FullyQualifiedErrorId |
                Should -Match '^EMPTY_SUBJECT'
        }
    }

    Context 'Quando o assunto possui prefixo de resposta' {

        It 'Deve rejeitar RE quando a permissÃ£o nÃ£o foi informada' {
            $subject = (
                'RE: RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject
            } | Should -Throw
        }

        It 'Deve aceitar RE quando AcceptReplyPrefixes estiver ativo' {
            $subject = (
                'RE: RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject `
                -AcceptReplyPrefixes

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'

            $result.ReplyPrefixRemoved |
                Should -Be $true
        }

        It 'Deve remover mÃºltiplos prefixos quando permitido' {
            $subject = (
                'ENC: RE: RelatÃ³rio de EliminaÃ§Ã£o Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject `
                -AcceptReplyPrefixes

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'

            $result.ReplyPrefixRemoved |
                Should -Be $true
        }
    }
}
