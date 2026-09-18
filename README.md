# Gerador PCP ESA

Aplicação web do Gerador PCP ESA, baseada na V9.8.7.

## Uso

A aplicação web fica em `index.html` e é preparada para GitHub Pages.

O navegador processa os PDFs ESA localmente usando PDF.js e gera os relatórios em PDF usando jsPDF + AutoTable.

### Funcionalidades

- Importação de um ou vários PDFs
- Detecção de múltiplas fases
- Manutenção da mesma O.S. em fases diferentes
- Dashboard PCP
- Relatório completo e resumido
- Consultas PCP
- Filtro de atrasadas por Dias na Fase ou Data de Entrega
- Consulta de O.S. selecionadas
- Gestão de atrasos por Fase + O.S.
- Motivo, observação, responsável, ação e prazo
- Histórico
- Colaborador
- Coluna OBS nos relatórios
- Armazenamento local no navegador

## Publicação no GitHub Pages

O workflow `.github/workflows/pages.yml` já está no repositório.

Na primeira publicação, abra:

**Settings → Pages → Source: GitHub Actions**

Depois disso, os próximos pushes em `main` publicam o site automaticamente.

A URL esperada é:

`https://frutiger-hub.github.io/Gerador-PCP-ESA/`

## Importante

Esta primeira versão é estática e salva os dados no navegador em uso. Isso significa que atrasos e histórico não são compartilhados automaticamente entre computadores.

Para dados centralizados entre vários usuários, será necessário adicionar um backend/banco externo sem deixar credenciais de escrita dentro do frontend.

## Dependências carregadas pelo navegador

- PDF.js 6.3.289
- jsPDF 4.2.1
- jsPDF AutoTable 5.0.8
