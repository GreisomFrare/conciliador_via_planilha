import os
import oracledb
from config_manager import load_config

_thick_initialized = False

DDL_CONCILIACAOCARTAO = """
CREATE TABLE CONCILIACAOCARTAO (
    ID_ARQUIVO         VARCHAR2(200) NOT NULL,
    DATA_PAGAMENTO     DATE,
    DATA_LANCAMENTO    DATE,
    ESTABELECIMENTO    VARCHAR2(30),
    TIPO_LANCAMENTO    VARCHAR2(100),
    FORMA_PAGAMENTO    VARCHAR2(100),
    BANDEIRA           VARCHAR2(50),
    VALOR_BRUTO        NUMBER(15,2),
    TAXA_TARIFA        NUMBER(15,2),
    VALOR_LIQUIDO      NUMBER(15,2),
    STATUS_PAGAMENTO   VARCHAR2(100),
    DATA_VENDA         DATE,
    HORA_VENDA         VARCHAR2(10),
    AUTORIZACAO        VARCHAR2(30),
    NSU_DOC            VARCHAR2(30),
    COD_VENDA          VARCHAR2(60),
    NUMERO_PARCELA     NUMBER(5),
    TOTAL_PARCELAS     NUMBER(5),
    TAXA_TOTAL         NUMBER(10,4),
    VALOR_TAXA_MDR     NUMBER(15,2),
    VALOR_TOTAL_TRANSACAO NUMBER(15,2),
    CONCILIADO         VARCHAR2(1) DEFAULT 'N'
)
"""

# Colunas retornadas pelo SQL do ERP (mesma ordem do SELECT)
COLS_ERP = [
    'ESTAB', 'IDACERFIN', 'IDACERCAR', 'NSUDOC', 'AUTORIZA',
    'IDCARTAO', 'PARCELAS', 'VALOR',
    'PARCELA', 'IDLANFIN',
    'VLRPARC', 'VLRTAXA', 'VLRL', 'TAXA',
    'DTVENCTO', 'DTBAIXA', 'DATACREDITO',
]


def _resolve_dir(path: str) -> str:
    if not path:
        return ''
    if os.path.isfile(path):
        return os.path.dirname(path)
    return path if os.path.isdir(path) else ''


def _init_thick_mode(lib_dir: str) -> bool:
    global _thick_initialized
    if _thick_initialized:
        return not oracledb.is_thin_mode()
    _thick_initialized = True
    try:
        if lib_dir and os.path.isdir(lib_dir):
            oracledb.init_oracle_client(lib_dir=lib_dir)
        else:
            oracledb.init_oracle_client()
        return True
    except Exception:
        return False


def get_connection(cfg: dict):
    alias = cfg.get('tns_alias', '').strip()
    if not alias:
        raise RuntimeError(
            'TNS Alias não configurado.\n'
            'Informe o Network Alias na aba Configuração.'
        )
    client_dir = _resolve_dir(cfg.get('oracle_client_dir', '').strip())
    thick_ok   = _init_thick_mode(client_dir)
    tns_dir    = _resolve_dir(cfg.get('tns_dir', '').strip())

    connect_kwargs = {
        'user':     cfg.get('user', ''),
        'password': cfg.get('password', ''),
        'dsn':      alias,
    }
    if thick_ok:
        if tns_dir:
            os.environ['TNS_ADMIN'] = tns_dir
    else:
        if tns_dir:
            connect_kwargs['config_dir'] = tns_dir

    return oracledb.connect(**connect_kwargs)


def _build_case_criteria(criteria: dict, de_para: dict, estab: str,
                         conciliado_only: bool = False) -> str:
    """Monta as condições do EXISTS dentro do CASE WHEN."""
    conditions = []
    if criteria.get('nsu'):
        conditions.append("CC.NSU_DOC = LTRIM(X.NSUDOC, '0')")
    if criteria.get('autorizacao'):
        conditions.append('CC.AUTORIZACAO = X.AUTORIZA')
    if criteria.get('data_venda'):
        conditions.append('CC.DATA_PAGAMENTO = A.DATA')
    if criteria.get('valor_bruto'):
        variacao = float(criteria.get('variacao') or 0)
        if variacao > 0:
            conditions.append(f'ABS(CC.VALOR_BRUTO - X1.VALOR) <= {variacao:.2f}')
        else:
            conditions.append('CC.VALOR_BRUTO = X1.VALOR')
    if criteria.get('estab') and de_para:
        cielo_codes = [c for c, e in de_para.items() if str(e) == str(estab)]
        if cielo_codes:
            quoted = ', '.join(f"'{c}'" for c in cielo_codes)
            conditions.append(f"CC.ESTABELECIMENTO IN ({quoted})")
        else:
            conditions.append('1=0')
    if criteria.get('parcela'):
        conditions.append('COALESCE(CC.NUMERO_PARCELA, 1) = COALESCE(X1.SEQPARC, 1)')
    if conciliado_only:
        conditions.append("CC.CONCILIADO = 'S'")
    # Se nenhum critério marcado, nunca dá match (evita UPDATE em tudo)
    return ' AND '.join(conditions) if conditions else '1=0'


class DatabaseManager:

    def __init__(self):
        self.connection = None

    def connect(self):
        self.connection = get_connection(load_config())

    def disconnect(self):
        if self.connection:
            try:
                self.connection.close()
            except Exception:
                pass
            self.connection = None

    def ensure_table_exists(self):
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                "SELECT COUNT(*) FROM USER_TABLES WHERE TABLE_NAME = 'CONCILIACAOCARTAO'"
            )
            if cursor.fetchone()[0] == 0:
                # Tabela não existe — cria
                cursor.execute(DDL_CONCILIACAOCARTAO)
                self.connection.commit()
            else:
                # Tabela existe — verifica se coluna foi renomeada em v4
                cursor.execute("""
                    SELECT COUNT(*) FROM USER_TAB_COLUMNS
                    WHERE TABLE_NAME = 'CONCILIACAOCARTAO'
                      AND COLUMN_NAME = 'NUM_PARCELA'
                """)
                if cursor.fetchone()[0] > 0:
                    cursor.execute(
                        "ALTER TABLE CONCILIACAOCARTAO "
                        "RENAME COLUMN NUM_PARCELA TO NUMERO_PARCELA"
                    )
                    self.connection.commit()
        finally:
            cursor.close()

    def arquivo_ja_importado(self, id_arquivo: str) -> bool:
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                'SELECT COUNT(*) FROM CONCILIACAOCARTAO WHERE ID_ARQUIVO = :1',
                [id_arquivo]
            )
            return cursor.fetchone()[0] > 0
        finally:
            cursor.close()

    def insert_conciliacao_records(self, id_arquivo: str, records: list):
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                'DELETE FROM CONCILIACAOCARTAO WHERE ID_ARQUIVO = :1',
                [id_arquivo]
            )
            sql = """
                INSERT INTO CONCILIACAOCARTAO (
                    ID_ARQUIVO, DATA_PAGAMENTO, DATA_LANCAMENTO, ESTABELECIMENTO,
                    TIPO_LANCAMENTO, FORMA_PAGAMENTO, BANDEIRA, VALOR_BRUTO,
                    TAXA_TARIFA, VALOR_LIQUIDO, STATUS_PAGAMENTO, DATA_VENDA,
                    HORA_VENDA, AUTORIZACAO, NSU_DOC, COD_VENDA,
                    NUMERO_PARCELA, TOTAL_PARCELAS, TAXA_TOTAL,
                    VALOR_TAXA_MDR, VALOR_TOTAL_TRANSACAO, CONCILIADO
                ) VALUES (
                    :id_arquivo, :data_pagamento, :data_lancamento, :estabelecimento,
                    :tipo_lancamento, :forma_pagamento, :bandeira, :valor_bruto,
                    :taxa_tarifa, :valor_liquido, :status_pagamento, :data_venda,
                    :hora_venda, :autorizacao, :nsu_doc, :cod_venda,
                    :numero_parcela, :total_parcelas, :taxa_total,
                    :valor_taxa_mdr, :valor_total_transacao, 'N'
                )
            """
            cursor.executemany(sql, records)
            self.connection.commit()
        finally:
            cursor.close()

    def get_records_by_arquivo(self, id_arquivo: str):
        cursor = self.connection.cursor()
        try:
            cursor.execute("""
                SELECT
                    CONCILIADO, DATA_VENDA, ESTABELECIMENTO, NSU_DOC,
                    AUTORIZACAO, VALOR_BRUTO, TAXA_TARIFA, VALOR_LIQUIDO,
                    BANDEIRA, NUMERO_PARCELA, TOTAL_PARCELAS,
                    FORMA_PAGAMENTO, STATUS_PAGAMENTO, DATA_PAGAMENTO
                FROM CONCILIACAOCARTAO
                WHERE ID_ARQUIVO = :1
                ORDER BY DATA_VENDA, NSU_DOC
            """, [id_arquivo])
            cols = [d[0] for d in cursor.description]
            rows = cursor.fetchall()
            return cols, rows
        finally:
            cursor.close()

    def get_acercar_records(self,
                            data_ini_str: str,
                            data_fim_str: str,
                            estab: str = None,
                            criteria: dict = None,
                            de_para: dict = None):
        sql = """
            SELECT
                X.ESTAB, X.IDACERFIN, X.IDACERCAR, X.NSUDOC, X.AUTORIZA,
                X.IDCARTAO, X.QTDEPARCELAS AS PARCELAS, X.VALOR AS VALOR,
                X1.SEQPARC AS PARCELA, X1.IDLANFIN,
                X1.VALOR AS VLRPARC, -X1.TXADM AS VLRTAXA,
                X1.VALOR - COALESCE(X1.TXADM, 0) AS VLRL,
                COALESCE(X1.TXADMP, 0) AS TAXA,
                X1.DTVENCTO, X1.DTBAIXA,
                A.DATA AS DATACREDITO
            FROM ACERCAR X
                INNER JOIN ACERCARDET X1
                    ON (X.ESTAB = X1.ESTAB AND X.IDACERFIN = X1.IDACERFIN AND X.IDACERCAR = X1.IDACERCAR)
                INNER JOIN ACERFIN A
                    ON (X1.ESTAB = A.ESTAB AND X1.IDACERFIN = A.IDACERFIN)
            WHERE A.DATA BETWEEN TO_DATE(:data_inicio, 'DD/MM/YYYY')
                             AND TO_DATE(:data_fim,    'DD/MM/YYYY')
            ORDER BY X.ESTAB, X1.SEQPARC, X.IDACERFIN
        """

        params = {
            'data_inicio': data_ini_str,
            'data_fim':    data_fim_str,
        }

        cursor = self.connection.cursor()
        try:
            cursor.execute(sql, params)
            rows = cursor.fetchall()
            return COLS_ERP, rows
        finally:
            cursor.close()

    def conciliar(self,
                  id_arquivo: str,
                  data_ini_str: str,
                  data_fim_str: str,
                  estab: str,
                  criteria: dict,
                  de_para: dict = None) -> int:
        de_para       = de_para or {}
        case_criteria = _build_case_criteria(criteria, de_para, estab)
        if case_criteria == '1=0':
            return 0

        sql = f"""
            UPDATE CONCILIACAOCARTAO CC
            SET CC.CONCILIADO = 'S'
            WHERE CC.ID_ARQUIVO = :id_arquivo
              AND CC.CONCILIADO = 'N'
              AND EXISTS (
                  SELECT 0
                  FROM ACERCAR X
                  INNER JOIN ACERCARDET X1
                      ON (X.ESTAB = X1.ESTAB
                          AND X.IDACERFIN = X1.IDACERFIN
                          AND X.IDACERCAR = X1.IDACERCAR)
                  INNER JOIN ACERFIN A
                      ON (X1.ESTAB = A.ESTAB
                          AND X1.IDACERFIN = A.IDACERFIN)
                  WHERE {case_criteria}
                    AND X.ESTAB = :estab
                    AND A.DATA BETWEEN TO_DATE(:data_inicio, 'DD/MM/YYYY')
                                   AND TO_DATE(:data_fim,    'DD/MM/YYYY')
              )
        """

        params = {
            'id_arquivo':  id_arquivo,
            'estab':       int(estab) if estab else 1003,
            'data_inicio': data_ini_str,
            'data_fim':    data_fim_str,
        }

        cursor = self.connection.cursor()
        try:
            cursor.execute(sql, params)
            n = cursor.rowcount
            self.connection.commit()
            return n
        finally:
            cursor.close()

    def debug_sql_conciliar(self,
                            id_arquivo: str,
                            data_ini_str: str,
                            data_fim_str: str,
                            estab: str,
                            criteria: dict,
                            de_para: dict = None) -> str:
        """Retorna o SQL de conciliação resolvido (para debug) sem executar."""
        de_para = de_para or {}
        case_criteria = _build_case_criteria(criteria, de_para, estab)
        estab_int = int(estab) if estab else 1003

        sql_resolved = f"""-- ===== DEBUG SQL CONCILIAÇÃO =====
-- Critérios ativos: {criteria}
-- De-Para: {de_para}
-- Estab config: {estab_int}
-- Case criteria montado: {case_criteria}

UPDATE CONCILIACAOCARTAO CC
SET CC.CONCILIADO = 'S'
WHERE CC.ID_ARQUIVO = '{id_arquivo}'
  AND CC.CONCILIADO = 'N'
  AND EXISTS (
      SELECT 0
      FROM ACERCAR X
      INNER JOIN ACERCARDET X1
          ON (X.ESTAB = X1.ESTAB
              AND X.IDACERFIN = X1.IDACERFIN
              AND X.IDACERCAR = X1.IDACERCAR)
      INNER JOIN ACERFIN A
          ON (X1.ESTAB = A.ESTAB
              AND X1.IDACERFIN = A.IDACERFIN)
      WHERE {case_criteria}
        AND X.ESTAB = {estab_int}
        AND A.DATA BETWEEN TO_DATE('{data_ini_str}', 'DD/MM/YYYY')
                       AND TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
  )"""
        return sql_resolved

    def desfazer_conciliacao(self, id_arquivo: str) -> int:
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                "UPDATE CONCILIACAOCARTAO SET CONCILIADO = 'N' WHERE ID_ARQUIVO = :1",
                [id_arquivo]
            )
            n = cursor.rowcount
            self.connection.commit()
            return n
        finally:
            cursor.close()
