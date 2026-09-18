# Gerador PCP ESA

Aplicação desktop Windows para leitura de relatórios PDF da ESA e geração de relatórios gerenciais de PCP.

## Conteúdo

- Aplicação V9.8.7
- Interface desktop em português
- Leitura de PDF
- Relatórios PCP
- Consultas de O.S.
- Gestão de atrasos
- Histórico
- Ícones e atalhos
- Build automática para Windows pelo GitHub Actions

## Executável

O workflow **Build Windows EXE** gera automaticamente `Gerador_PCP_ESA.exe` como artefato do GitHub Actions.

## Dependências

O build usa Python 3.14, PyMuPDF e ReportLab. O usuário final não precisa instalar Python quando usar o EXE gerado.

> O código-fonte e os recursos da versão 9.8.7 devem estar em `Aplicacao/` para o workflow produzir o EXE.
