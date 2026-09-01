"""
mapper.py — Lógica de leitura, mapeamento e escrita de planilhas Excel.
"""

from typing import Callable, Dict, List, Optional, Tuple
import openpyxl


def get_columns(filepath: str, sheet_index: int = 0) -> Tuple[List[str], str]:
    """
    Lê os cabeçalhos (primeira linha) de uma planilha Excel.

    Args:
        filepath: Caminho para o arquivo .xlsx
        sheet_index: Índice da aba (0 = primeira aba)

    Returns:
        Tupla (lista_de_colunas, nome_da_aba)
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    ws = wb[sheet_names[sheet_index]]
    headers = []
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        for cell in row:
            headers.append(str(cell) if cell is not None else "")
    wb.close()
    return headers, sheet_names[sheet_index]


def get_sheet_names(filepath: str) -> List[str]:
    """Retorna os nomes de todas as abas de uma planilha."""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def execute_mapping(
    source_path: str,
    dest_path: str,
    output_path: str,
    mapping: List[Dict],
    source_sheet: int = 0,
    dest_sheet: int = 0,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> int:
    """
    Aplica o mapeamento de colunas: copia dados da planilha de origem
    para a planilha de destino e salva em output_path.

    Args:
        source_path: Caminho da planilha de origem (fonte dos dados)
        dest_path: Caminho da planilha de destino (template)
        output_path: Onde salvar o arquivo preenchido
        mapping: Lista de {'source': 'NomeColOrigem', 'destination': 'NomeColDestino'}
        source_sheet: Índice da aba na planilha de origem
        dest_sheet: Índice da aba na planilha de destino
        progress_callback: Função chamada com (linha_atual, total_linhas)

    Returns:
        Número de linhas copiadas.
    """
    # Lê origem
    wb_src = openpyxl.load_workbook(source_path, data_only=True)
    ws_src = wb_src[wb_src.sheetnames[source_sheet]]

    # Lê destino (template)
    wb_dst = openpyxl.load_workbook(dest_path)
    ws_dst = wb_dst[wb_dst.sheetnames[dest_sheet]]

    # Mapeia nomes de colunas para índices na origem
    src_headers = [str(c.value) if c.value is not None else "" for c in ws_src[1]]
    dst_headers = [str(c.value) if c.value is not None else "" for c in ws_dst[1]]

    # Cria dicionário: coluna_destino -> coluna_origem (índice 1-based)
    col_map: Dict[int, int] = {}  # dst_col_idx -> src_col_idx
    for rule in mapping:
        src_name = rule.get("source", "")
        dst_name = rule.get("destination", "")
        if src_name in src_headers and dst_name in dst_headers:
            src_idx = src_headers.index(src_name) + 1
            dst_idx = dst_headers.index(dst_name) + 1
            col_map[dst_idx] = src_idx

    # Conta linhas de dados na origem (sem o cabeçalho)
    src_rows = list(ws_src.iter_rows(min_row=2, values_only=True))
    total = len(src_rows)

    # Determina a primeira linha de dados no destino
    # Preserva as linhas de cabeçalho do template (linha 1)
    dst_data_start = 2

    # Limpa dados existentes no destino a partir da linha de dados
    # (mantém o cabeçalho intacto)
    if ws_dst.max_row >= dst_data_start:
        for row in ws_dst.iter_rows(min_row=dst_data_start):
            for cell in row:
                cell.value = None

    # Copia dados linha a linha
    for i, src_row_values in enumerate(src_rows):
        dst_row_num = dst_data_start + i

        for dst_col_idx, src_col_idx in col_map.items():
            # src_row_values é 0-indexed
            value = src_row_values[src_col_idx - 1]
            ws_dst.cell(row=dst_row_num, column=dst_col_idx, value=value)

        if progress_callback:
            progress_callback(i + 1, total)

    wb_dst.save(output_path)
    wb_src.close()

    return total
