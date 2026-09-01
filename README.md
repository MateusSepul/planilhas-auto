# 🗂 Mapeador de Planilhas

Ferramenta desktop com interface gráfica para **copiar dados entre planilhas Excel**, mapeando colunas de origem para colunas de destino de forma visual e intuitiva. Suporta salvar e reutilizar perfis de mapeamento (layouts).

---

## ✨ Funcionalidades

- 📂 Seleção visual de planilhas de **origem** e **destino** (`.xlsx` / `.xls`)
- 🔀 **Mapeamento de colunas** por clique: selecione uma coluna de origem → selecione uma coluna de destino
- 👁 Visualização em tempo real dos mapeamentos ativos
- 💾 **Salvar e reutilizar layouts** de mapeamento (perfis nomeados)
- 🗑 Gerenciamento de layouts: salvar, carregar e deletar
- ▶ Execução do preenchimento em **thread separada** com barra de progresso
- 🎨 Interface escura moderna (tema dark com paleta roxa/verde)

---

## 🏗 Arquitetura do Projeto

```
planilhas-auto/
├── main.py           # Ponto de entrada — inicia a aplicação
├── app.py            # Interface gráfica principal (Tkinter)
├── mapper.py         # Lógica de leitura e escrita de planilhas (openpyxl)
├── layout_store.py   # Persistência de layouts em JSON
├── layouts.json      # Arquivo de layouts salvos
└── requirements.txt  # Dependências Python
```

### Responsabilidades por módulo

| Módulo | Responsabilidade |
|---|---|
| `main.py` | Ponto de entrada; instancia e executa o `App` |
| `app.py` | Toda a UI com Tkinter: janela, painéis, eventos, estilos |
| `mapper.py` | Lê cabeçalhos, executa cópia de dados entre planilhas |
| `layout_store.py` | CRUD de layouts no arquivo `layouts.json` |

---

## 🔄 Fluxo de Uso

```
1. Abrir planilha de ORIGEM
       ↓
2. Abrir planilha de DESTINO (template)
       ↓
3. Mapear colunas:
   clique na coluna de Origem → clique na coluna de Destino
   (repita para cada par desejado)
       ↓
4. (Opcional) Salvar o mapeamento como Layout nomeado
       ↓
5. Clicar em "▶ Preencher Planilha"
       ↓
6. Escolher onde salvar o arquivo de saída (.xlsx)
       ↓
7. ✅ Planilha gerada com os dados copiados
```

---

## ⚙️ Como Funciona Internamente

### 1. Leitura das planilhas (`mapper.py`)

- `get_sheet_names(filepath)` — lista as abas disponíveis no arquivo Excel
- `get_columns(filepath, sheet_index)` — lê a **primeira linha** da aba escolhida como cabeçalhos

### 2. Mapeamento de colunas (`app.py`)

O estado de mapeamento é um dicionário:
```python
self.mappings: Dict[str, str] = {
    "coluna_destino": "coluna_origem",
    ...
}
```

O usuário constrói esse dicionário clicando nas listas visuais. Cada par fica visível no painel **"Mapeamentos Ativos"** com botão de remoção individual.

### 3. Execução do preenchimento (`mapper.py → execute_mapping`)

```
Origem (.xlsx)          Destino/Template (.xlsx)
┌──────────────┐        ┌──────────────────────┐
│ Nome │ Email │  ───▶  │ Cliente │ Contato │..│
│ João │ j@... │        │ João    │ j@...   │  │
│ ...  │ ...   │        │ ...     │ ...     │  │
└──────────────┘        └──────────────────────┘
                                 ↓
                        Salvo em output.xlsx
```

Etapas internas:
1. Carrega a planilha de origem (`openpyxl`, modo `read_only`)
2. Carrega o template de destino (`openpyxl`, modo leitura/escrita)
3. Resolve nomes de colunas para índices numéricos em ambas as planilhas
4. **Apaga os dados existentes** no destino (mantém o cabeçalho da linha 1)
5. Copia linha a linha os valores mapeados da origem para o destino
6. Chama `progress_callback(linha_atual, total)` a cada linha para atualizar a barra
7. Salva o arquivo resultante em `output_path`

### 4. Layouts (`layout_store.py`)

Os layouts são salvos em `layouts.json` no mesmo diretório do projeto:

```json
{
  "Layout Vendas": [
    { "source": "Nome",  "destination": "Cliente" },
    { "source": "Valor", "destination": "Total"   }
  ]
}
```

Operações disponíveis: `save_layout`, `load_layout`, `list_layouts`, `delete_layout`.

### 5. Thread de execução (`app.py → _execute`)

Para não travar a interface durante o processamento, a cópia roda em uma **thread separada** (`threading.Thread(daemon=True)`). A barra de progresso e o label de status são atualizados via `self.update_idletasks()` a cada linha processada.

---

## 🚀 Instalação e Execução

### Pré-requisitos

- Python 3.8+
- `tkinter` (incluso na instalação padrão do Python no Windows)

### Instalar dependências

```bash
pip install -r requirements.txt
```

> A única dependência externa é o `openpyxl >= 3.1.0`.

### Executar

```bash
python main.py
```

---

## 📋 Exemplo de Uso

Você tem dois arquivos:

**`origem.xlsx`** — dados brutos exportados de um sistema:
| Nome | Email | Valor | Data |
|------|-------|-------|------|
| João | j@x.com | 100 | 2024-01-01 |

**`template.xlsx`** — planilha de destino com cabeçalhos próprios:
| Cliente | Contato | Total | Ref |
|---------|---------|-------|-----|

Ao mapear `Nome → Cliente`, `Email → Contato`, `Valor → Total`, `Data → Ref` e executar, o resultado será:

**`saida.xlsx`**:
| Cliente | Contato | Total | Ref |
|---------|---------|-------|-----|
| João | j@x.com | 100 | 2024-01-01 |

---

## 🗂 Estrutura do `layouts.json`

```json
{
  "Nome do Layout": [
    { "source": "NomeColOrigem", "destination": "NomeColDestino" },
    ...
  ]
}
```

O arquivo é criado automaticamente ao salvar o primeiro layout e fica no mesmo diretório do projeto.

---

## 🛠 Tecnologias

| Tecnologia | Uso |
|---|---|
| Python 3 | Linguagem principal |
| Tkinter | Interface gráfica nativa (sem dependências externas) |
| openpyxl | Leitura e escrita de arquivos `.xlsx` |
| threading | Processamento assíncrono sem travar a UI |
| JSON | Persistência dos layouts de mapeamento |
