# Conciliador de Cartão via Planilha — Documentação Base

> **Manter este arquivo atualizado a cada alteração no código.**

---

## O que é este projeto

Aplicação desktop (Windows) que concilia transações de cartão de crédito da **Cielo** com lançamentos contábeis do ERP **ACERCAR/VIASUPER**. O usuário importa um arquivo Excel exportado da Cielo, busca os lançamentos correspondentes no ERP via Oracle, e executa a conciliação (marca registros como conciliados).

---

## Stack

| Tecnologia | Uso |
|---|---|
| Python 3.x + Tkinter | Interface gráfica desktop |
| oracledb 2.0+ (thick mode) | Conexão com banco Oracle via TNS |
| openpyxl | Leitura de arquivos Excel (.xlsx) |
| PyInstaller | Empacotamento em .exe único |
| JSON | Persistência de configurações locais |

---

## Estrutura de arquivos

```
conciliador_via_planilha/
├── main.py              # Bootstrap: inicializa janela Tk e inicia o app
├── app.py               # Toda a GUI (1164 linhas, 3 abas, 2 grids)
├── config_manager.py    # Leitura/gravação de config.json
├── database.py          # Todas as operações Oracle (DDL, queries, UPDATE)
├── excel_importer.py    # Parser do Excel da Cielo
├── requirements.txt     # Dependências Python
├── build.bat            # Script PyInstaller para gerar o .exe
└── config.json          # Gerado em runtime — NÃO está no git (credenciais)
```

---

## Módulos em detalhe

### `main.py`
Ponto de entrada. Detecta se está rodando como `.exe` (PyInstaller frozen) ou script Python para ajustar o `sys.path`. Cria a janela `Tk()`, instancia `ConciliacaoCartaoApp` e chama `mainloop()`.

---

### `config_manager.py`
Gerencia o arquivo `config.json` (criado ao lado do `.exe` ou do script).

**Estrutura do config.json:**
```json
{
  "tns_alias": "",
  "user": "",
  "password": "",
  "tns_dir": "",
  "oracle_client_dir": "",
  "estab": "1003",
  "de_para": [
    ["1016619208", "1003"],
    ["1098340331", "1010"],
    ["1098940242", "1011"],
    ["2786503480", "1012"],
    ["2786503498", "1013"]
  ]
}
```

- `de_para`: lista de pares `[código_cielo, código_erp]` — mapeamento de estabelecimentos.
- Funções: `load_config()` → retorna dict com defaults; `save_config(cfg)` → persiste.

---

### `excel_importer.py`
Lê o arquivo Excel exportado da Cielo.

**Fluxo do `import_excel(file_path)`:**
1. Nome do arquivo (sem extensão) vira o `id_arquivo` — chave primária do lote.
2. Abre o workbook com `data_only=True`.
3. Localiza a linha de cabeçalho procurando `'Data de pagamento'` nas primeiras 25 linhas.
4. Detecta automaticamente as colunas de parcela pelo cabeçalho (fuzzy match); usa índices padrão como fallback.
5. Itera as linhas, converte tipos:
   - Datas: aceita `DD/MM/YYYY`, `YYYY-MM-DD`, `MM/DD/YYYY` e objetos `date`.
   - Decimais: aceita vírgula como separador decimal (formato BR).
   - Inteiros: aceita float e string.
6. Retorna `(id_arquivo, lista_de_dicts)`.

**Constantes de índice de coluna** (fallback quando header não é encontrado): definidas nas linhas 10–29.

---

### `database.py`
Toda a camada de dados. Usa **thick mode** do oracledb (requer Oracle Client instalado).

#### Tabela `CONCILIACAOCARTAO`
Criada automaticamente se não existir. Armazena os registros importados da Cielo.

Colunas relevantes:
| Coluna | Tipo | Descrição |
|---|---|---|
| ID_ARQUIVO | VARCHAR2(200) | Nome do arquivo Excel (chave do lote) |
| CONCILIADO | CHAR(1) | 'S' = conciliado, 'N' = pendente |
| NSU_DOC | VARCHAR2(50) | NSU/documento da transação |
| AUTORIZACAO | VARCHAR2(50) | Código de autorização |
| DATA_PAGAMENTO | DATE | Data de crédito/pagamento |
| VALOR_BRUTO | NUMBER | Valor bruto da transação |
| ESTABELECIMENTO | VARCHAR2(50) | Código do estabelecimento Cielo |
| NUMERO_PARCELA | NUMBER | Número da parcela |

> **Migração automática:** se a coluna `NUM_PARCELA` existir (versão antiga), ela é renomeada para `NUMERO_PARCELA` no `ensure_table_exists()`.

#### Classe `DatabaseManager`
| Método | O que faz |
|---|---|
| `connect()` | Abre conexão Oracle via TNS; configura `TNS_ADMIN` env var |
| `disconnect()` | Fecha conexão |
| `ensure_table_exists()` | Cria tabela se não existe; aplica migração de coluna |
| `arquivo_ja_importado(id)` | Retorna True se o lote já existe |
| `insert_conciliacao_records(id, records)` | DELETE + INSERT em batch |
| `get_records_by_arquivo(id)` | Busca todos os registros do lote (14 colunas) |
| `get_acercar_records(...)` | Query no ERP: join ACERCAR → ACERCARDET → ACERFIN |
| `conciliar(...)` | UPDATE com EXISTS subquery — marca CONCILIADO = 'S' |
| `debug_sql_conciliar(...)` | Retorna o SQL gerado sem executar |
| `desfazer_conciliacao(id)` | Reverte todos os CONCILIADO para 'N' do lote |

#### Query ERP (`get_acercar_records`)
Join triplo: `ACERCAR` (cabeçalho) → `ACERCARDET` (parcelas) → `ACERFIN` (financeiro/crédito).
Filtro principal: `ACERFIN.DATA BETWEEN :data_ini AND :data_fim`.

#### Filtro de estabelecimento (`_build_estab_filter`)
Determina como o `EXISTS` filtra o ERP pelo estabelecimento:

- **Com de-para preenchido**: gera um `CASE CC.ESTABELECIMENTO WHEN cielo THEN erp ...` — cada registro Cielo resolve automaticamente para seu ERP correto, suportando múltiplos estabelecimentos.
- **Sem de-para**: fallback para `X.ESTAB = :estab` (usa o estab único configurado).

Exemplo com de-para:
```sql
X.ESTAB = CASE CC.ESTABELECIMENTO
    WHEN '1016619208' THEN 1003
    WHEN '1098340331' THEN 1010
    WHEN '1098940242' THEN 1011
    ELSE NULL
END
```

> **Por que CASE e não IN?** O `IN` seria mais permissivo (qualquer Cielo poderia bater com qualquer ERP). O `CASE` garante que cada Cielo só bate com o ERP correspondente ao seu mapeamento.

#### Lógica de conciliação (`_build_case_criteria`)
Monta condições `WHERE` baseadas nos checkboxes do usuário:

| Checkbox | Condição SQL gerada |
|---|---|
| NSU/DOC | `CC.NSU_DOC = LTRIM(X.NSUDOC, '0')` |
| Autorização | `CC.AUTORIZACAO = X.AUTORIZA` |
| Data Pgto = Data Crédito | `CC.DATA_PAGAMENTO = A.DATA` |
| Valor (com tolerância) | `ABS(CC.VALOR_BRUTO - X1.VALOR) <= :tolerancia` |
| Estabelecimento (de-para) | `CC.ESTABELECIMENTO IN (códigos_cielo_do_estab)` |
| Parcela | `COALESCE(CC.NUMERO_PARCELA,1) = COALESCE(X1.SEQPARC,1)` |

Se nenhum critério for selecionado, a condição `1=0` impede qualquer conciliação.

> **NSU/DOC — zeros à esquerda**: o ERP armazena o NSU com zeros à esquerda (ex: `000001`); o `LTRIM(X.NSUDOC, '0')` remove esses zeros antes de comparar com o valor Cielo.

---

### `app.py`
Interface gráfica. Classe principal: `ConciliacaoCartaoApp`.

#### Classes auxiliares internas
- **`_CalendarPopup`**: Widget de calendário custom em PT-BR, com navegação mês/ano, aparece abaixo do campo de data.
- **`_DateEntry`**: Campo de data composto (entrada texto + botão calendário). Valida formato `DD/MM/YYYY`, pinta de vermelho se inválido.

#### Funções de formatação (topo do arquivo)
| Função | Converte |
|---|---|
| `_fmt(value)` | date → `DD/MM/YYYY`; None → `""` |
| `_fmt_num(value)` | float → `"1.234,56"` (formato BR) |
| `_br_to_float(s)` | `"1.234,56"` → `1234.56` |
| `_sort_key(val)` | Retorna tupla para ordenação multi-tipo (número < data < texto) |

#### Aba 1 — Principal
- Seleção e importação do arquivo Excel.
- Filtros de data (início/fim) com calendário.
- Botões: **Buscar Lançamentos** (ERP), **Conciliar**, **Desfazer**, **Debug SQL**.
- 6 checkboxes de critério + campo de tolerância de valor.
- **Grid esquerdo**: dados Cielo (lote importado).
- **Grid direito**: dados ERP (ACERCAR).
- Linha de filtros por coluna acima de cada grid.
- Cabeçalho clicável para ordenação (▲/▼).
- Scroll horizontal sincronizado entre filtros e grid.
- Cores: verde = conciliado, amarelo = pendente, azul claro = matched.
- Totalizadores de registros e valores.

**Colunas do Grid Esquerdo (Cielo):**

| Índice | Nome interno | Header exibido |
|---|---|---|
| 0 | conciliado | OK? |
| 1 | data_venda | Data Venda |
| 2 | estabelecimento | Estabelecimento |
| 3 | estab_erp | Estab ERP |
| 4 | nsu_doc | NSU/DOC |
| 5 | autorizacao | Autorização |
| 6 | valor_bruto | Valor Bruto |
| 7 | valor_liquido | Valor Líquido |
| 8 | bandeira | Bandeira |
| 9 | parcela | Parc. |
| 10 | total_parc | Total Parc. |
| 11 | forma_pag | Forma Pgto |
| 12 | status_pag | Status |

**Colunas do Grid Direito (ERP):**

| Índice | Nome interno | Header exibido | Numérico? |
|---|---|---|---|
| 0 | existenoarq | OK? | Não |
| 1 | estab | Estab. | Não |
| 2 | nsudoc | NSU/DOC | Não |
| 3 | autoriza | Autorização | Não |
| 4 | idcartao | Cartão | Não |
| 5 | parcela | Parc. | Não |
| 6 | parcelas | Total Parc. | Não |
| 7 | vlrparc | Vlr Parcela | **Sim** |
| 8 | vlrtaxa | Taxa | **Sim** |
| 9 | vlrl | Vlr Líquido | **Sim** |
| 10 | taxa | % Taxa | **Sim** |
| 11 | dtvencto | Dt. Vencto | Não |
| 12 | datacredito | Dt. Crédito | Não |

#### Aba 2 — De-Para Estabelecimento
Mapeamento de código Cielo → código ERP. Os pares são exibidos em Treeview, persistidos no `config.json`. Botões: Adicionar, Remover, Salvar.

#### Aba 3 — Configuração Oracle
Campos: TNS Alias, Usuário, Senha, Código do Estabelecimento, Caminho do `tnsnames.ora`, Diretório do Oracle Client. Botão "Testar Conexão". Tudo salvo em `config.json`.

---

## Fluxos principais

### Importar arquivo Excel
```
Seleciona arquivo
    ↓
_cmd_importar() [thread]
    ↓
import_excel() → (id_arquivo, records)
    ↓
db.connect()
db.ensure_table_exists()
db.arquivo_ja_importado() → pergunta se reimportar
db.insert_conciliacao_records()   ← DELETE + INSERT
    ↓
UI: mensagem de sucesso / erro
db.disconnect()
```

### Buscar lançamentos ERP
```
Define período + critérios
Clica "Buscar Lançamentos"
    ↓
_cmd_buscar() [thread]
    ↓
db.get_acercar_records(datas, estab, criteria, de_para)
    └─ JOIN ACERCAR / ACERCARDET / ACERFIN
    ↓
Formata valores (BR) → popula grid direito
Atualiza totalizadores
db.disconnect()
```

### Conciliar
```
Clica "Conciliar"
    ↓
Confirmação dialog
    ↓
_cmd_conciliar() [thread]
    ↓
_build_case_criteria() → condições WHERE
db.conciliar()
    └─ UPDATE CONCILIACAOCARTAO SET CONCILIADO='S'
       WHERE EXISTS (subquery ERP com critérios)
    └─ Retorna count de linhas atualizadas
    ↓
_reload_both_grids()   ← recarrega ambos os grids
Mensagem com total conciliado
db.disconnect()
```

### Desfazer conciliação
```
Clica "Desfazer"
    ↓
Confirmação dialog
    ↓
db.desfazer_conciliacao(id_arquivo)
    └─ UPDATE SET CONCILIADO='N' WHERE ID_ARQUIVO = :id
    ↓
_reload_both_grids()
```

---

## Threading

Operações longas rodam em `threading.Thread(daemon=True)` para não travar a UI:
- Importação, Busca ERP, Conciliação, Desfazer, Teste de conexão.

Atualizações de UI dentro de threads usam `self.root.after(0, func)` (thread-safe no Tkinter).

---

## Build / Distribuição

`build.bat` executa PyInstaller:
- `--onefile`: gera um único `.exe`
- `--noconsole`: sem janela de console
- `--name "ConciliacaoCartao"`: nome do executável
- `--hidden-import`: inclui `oracledb`, `openpyxl` e submódulos
- `--add-data`: embute os módulos Python no executável

Saída em `dist/ConciliacaoCartao.exe`.

---

## Pontos de atenção para manutenção

1. **Thick mode Oracle**: requer Oracle Instant Client instalado na máquina. O caminho é configurado pelo usuário na Aba 3.
2. **TNS_ADMIN**: o `database.py` seta a variável de ambiente `TNS_ADMIN` antes de conectar para apontar ao `tnsnames.ora` configurado.
3. **id_arquivo**: é o nome do arquivo Excel sem extensão. É a chave do lote — renomear o arquivo muda o id e reimporta como novo lote.
4. **Migração de schema**: `ensure_table_exists()` trata a renomeação `NUM_PARCELA → NUMERO_PARCELA`. Se houver novas mudanças de schema, adicionar lógica ali.
5. **De-para**: sem mapeamento correto, o critério "Estabelecimento" não funciona. Os defaults cobrem 5 estabelecimentos (1003, 1010, 1011, 1012, 1013).
6. **Tolerância de valor**: padrão 0.05 (R$ 0,05). Controla o checkbox de valor.

---

## Histórico de alterações

| Data | Descrição |
|---|---|
| 2026-05-08 | Documentação base criada a partir da implementação inicial |
| 2026-05-08 | Correção: `conciliar()` e `debug_sql_conciliar()` agora usam CASE dinâmico do de-para para filtro de estab, suportando múltiplos estabelecimentos simultaneamente |
| 2026-05-08 | Melhoria: grid ERP agora mostra status real de conciliação via coluna EXISTENOARQ (EXISTS corr. em CONCILIACAOCARTAO com mesmos critérios do UPDATE). `_SHOW_DIR[0]` → índice 17. Contadores e cores do lado ERP agora funcionam corretamente |
| 2026-05-08 | Correção: critério "Estab ERP = Estab." em `_build_case_criteria` agora usa todos os códigos Cielo do de-para (não apenas os do estab configurado). O filtro de estab por registro já é feito pelo `_build_estab_filter` |
| 2026-05-08 | Correção: critério de data em `_build_case_criteria` corrigido de `CC.DATA_PAGAMENTO = A.DATA` para `CC.DATA_VENDA = A.DATA` — é a data da venda (transação) que corresponde à data de crédito do ERP, não a data de pagamento. Label do checkbox corrigido de "Dt. Pagamento" para "Dt. Venda" |
| 2026-05-08 | Otimização: `get_acercar_records` agora usa `A.DATA IN (SELECT DISTINCT DATA_VENDA FROM CONCILIACAOCARTAO WHERE ID_ARQUIVO = :id)` quando há arquivo carregado, em vez de BETWEEN amplo. Reduz drasticamente o volume retornado (~400K → apenas registros das datas presentes na planilha). Fallback para BETWEEN quando nenhum arquivo está carregado. |
| 2026-05-08 | Funcionalidade: edição inline de NSU/DOC e Autorização no grid ERP (duplo clique) somente quando o campo está nulo/vazio no ERP. Botão "Salvar Alterações ERP" executa UPDATE em ACERCAR com proteção via WHERE (só atualiza campos nulos/vazios). Novo método `update_acercar_nsu_autoriza()` em database.py. |
