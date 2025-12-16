# Lista de Tarefas por Desenvolvedor

Este projeto é um frontend simples que importa um arquivo Excel (`.xlsx`), extrai tarefas da aba semanal mais recente, filtra por desenvolvedor e gera uma lista de tarefas persistida localmente no navegador.

---

## Visão Geral

- Fonte de dados: planilha Excel com múltiplas abas
- Cada aba semanal representa um snapshot de tarefas
- Apenas a aba mais recente é utilizada
- O usuário escolhe um desenvolvedor e importa somente suas tarefas
- Estado da aplicação é salvo em `localStorage`

---

## Formato do Arquivo Excel

### Tipo
- `.xlsx`

### Estrutura
- Contém múltiplas abas
- Existem abas fixas que devem ser ignoradas (ex: `Matriz`, `SQL`, `Regras`)

---

## Abas Válidas

- Nome das abas no formato: `DD-MM`
  - Exemplos: `05-12`, `12-12`, `19-12`
- Cada aba válida representa um snapshot semanal
- Sempre deve ser utilizada a aba mais recente cronologicamente

---

## Estrutura da Aba Semanal

Cada linha da aba representa uma tarefa.

### Colunas Relevantes

| Coluna      | Tipo   | Descrição |
|-------------|--------|-----------|
| Tarefa      | number | Identificador único da tarefa |
| Título      | string | Título ou descrição curta |
| Cor         | string | Opcional, pode estar vazia |
| Dev Atual   | string | Nome do desenvolvedor responsável |

> Todas as demais colunas devem ser ignoradas pela aplicação.

---

## Regras de Importação

1. O usuário importa um arquivo `.xlsx`
2. A aplicação identifica as abas no formato `DD-MM`
3. A aba mais recente é selecionada
4. Todas as linhas da aba são lidas
5. Valores únicos da coluna **Dev Atual** são extraídos
6. O usuário seleciona um desenvolvedor
7. Apenas tarefas com `Dev Atual` exatamente igual ao valor selecionado são importadas

---

## Modelo de Dados Interno

As tarefas importadas são convertidas para o seguinte modelo interno:

```txt
Task
- id: number        // vindo da coluna "Tarefa"
- title: string     // vindo da coluna "Título"
- color: string     // vindo da coluna "Cor" (fallback se vazio)
- completed: boolean
- notes: string
