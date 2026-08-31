BeforeAll {
    $modulePath = Join-Path `
        $PSScriptRoot `
        '..\src\SubjectParser.psm1'

    Import-Module $modulePath -Force
}

Describe 'ConvertFrom-PatrimonialSubject' {

    Context 'Quando o assunto é válido' {

        It 'Deve extrair corretamente todas as informações' {
            $subject = (
                'Relatório de Eliminação Patrimonial - ' +
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

        It 'Deve normalizar espaços adicionais' {
            $subject = (
                '  Relatório de Eliminação Patrimonial  -  ' +
                'TBRA  (0185)  X  ITIS (0306)  -  08/2026  '
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'

            $result.Competence |
                Should -Be '08/2026'
        }

        It 'Deve converter as siglas para letras maiúsculas' {
            $subject = (
                'Relatório de Eliminação Patrimonial - ' +
                'tbra (0185) X itis (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject

            $result.Company1Acronym | Should -Be 'TBRA'
            $result.Company2Acronym | Should -Be 'ITIS'
        }
    }

    Context 'Quando o assunto é inválido' {

        It 'Deve lançar erro quando o formato está incorreto' {
            $subject = 'Relatório incorreto - TBRA X ITIS'

            {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject
            } | Should -Throw
        }

        It 'Deve retornar INVALID_SUBJECT_FORMAT para assunto incorreto' {
            $subject = 'Relatório incorreto - TBRA X ITIS'

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
                Should -Match 'não corresponde ao padrão esperado'
        }

        It 'Deve rejeitar mês maior que doze' {
            $subject = (
                'Relatório de Eliminação Patrimonial - ' +
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

        It 'Deve rejeitar código com menos de quatro dígitos' {
            $subject = (
                'Relatório de Eliminação Patrimonial - ' +
                'TBRA (185) X ITIS (0306) - 08/2026'
            )

            {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject
            } | Should -Throw
        }

        It 'Deve retornar EMPTY_SUBJECT para assunto com espaços' {
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

        It 'Deve rejeitar RE quando a permissão não foi informada' {
            $subject = (
                'RE: Relatório de Eliminação Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            {
                ConvertFrom-PatrimonialSubject `
                    -Subject $subject
            } | Should -Throw
        }

        It 'Deve aceitar RE quando AcceptReplyPrefixes estiver ativo' {
            $subject = (
                'RE: Relatório de Eliminação Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject `
                -AcceptReplyPrefixes

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'

            $result.ReplyPrefixRemoved |
                Should -BeTrue
        }

        It 'Deve remover múltiplos prefixos quando permitido' {
            $subject = (
                'ENC: RE: Relatório de Eliminação Patrimonial - ' +
                'TBRA (0185) X ITIS (0306) - 08/2026'
            )

            $result = ConvertFrom-PatrimonialSubject `
                -Subject $subject `
                -AcceptReplyPrefixes

            $result.RouteKey |
                Should -Be '0185:TBRA|0306:ITIS'

            $result.ReplyPrefixRemoved |
                Should -BeTrue
        }
    }
}
