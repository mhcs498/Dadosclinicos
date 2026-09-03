# 🏥 Dashboard Clínico - Ficha de Registo e Relatórios

Uma aplicação web leve, segura e independente de servidor (offline-first) para gestão de fichas de registo clínico, visualização de estatísticas em tempo real e exportação de relatórios profissionais.

## ✨ Funcionalidades Principais

- **📝 Ficha de Registo Completa:** ID automático, cálculo automático de faixa etária (dias/meses/anos), e campo de diagnóstico totalmente livre.
- **✏️ Tabela Editável Inline:** Todos os campos da tabela podem ser editados diretamente com um clique. Preparada para até 200 registos.
- **📊 Gráficos Interativos:** Distribuição por Sexo, Destino, Proveniência, Faixa Etária e Diagnóstico (com frequências absolutas e percentagens).
- **📤 Múltiplas Opções de Exportação:**
  - **Excel (.xlsx):** Ficheiro com múltiplas folhas (Registos, Estatísticas e Tabelas Dinâmicas com filtros).
  - **Word (.doc):** Relatório formatado com tabelas e gráficos embutidos.
  - **Imagens (PNG):** Exportação individual ou em lote de todos os gráficos.
  - **JSON:** Backup completo dos dados.
- **🔗 Partilha entre Dispositivos:** Sistema inteligente para partilhar os dados via Link ou Ficheiro HTML autónomo, sem necessidade de base de dados externa.

---

## 🚀 Como Usar

### 1. Execução Local
Basta fazer o download do ficheiro `Qwen_html_20260903_MAREKA.html` e abri-lo com qualquer navegador moderno (Chrome, Edge, Firefox, Safari). Não requer instalação nem servidor.

### 2. Adicionar Registos
Preencha o formulário no topo da página e clique em "➕ Adicionar Registo". Os dados são guardados automaticamente no navegador (`localStorage`).

### 3. Editar Dados
Clique diretamente em qualquer célula da tabela para alterar um valor. Os gráficos e estatísticas atualizam-se em tempo real.

---

## 🔗 Como Partilhar Dados entre Dispositivos

Como esta aplicação não usa uma base de dados centralizada, existem duas formas principais de partilhar os dados com outros computadores ou telemóveis:

### Método 1: Link Direto (Recomendado e Imediato)
O sistema comprime os dados e insere-os no final do URL. Pode usar o link "Raw" do GitHub para partilhar:

```text
https://raw.githubusercontent.com/mhcs498/dashboard--HCN/main/Qwen_html_20260903_MAREKA.html#d=[CODIGO_COMPIMIDO_AQUI]
