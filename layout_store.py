"""
layout_store.py — Persistência de layouts de mapeamento em JSON.
"""

import json
import os
from typing import Dict, List, Optional

LAYOUTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "layouts.json")


def _load_all() -> Dict:
    if not os.path.exists(LAYOUTS_FILE):
        return {}
    with open(LAYOUTS_FILE, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return {}


def _save_all(data: Dict) -> None:
    with open(LAYOUTS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def list_layouts() -> List[str]:
    """Retorna lista de nomes de layouts salvos."""
    return list(_load_all().keys())


def save_layout(name: str, mapping: List[Dict]) -> None:
    """
    Salva um layout de mapeamento.

    Args:
        name: Nome do perfil (ex: 'Layout Vendas')
        mapping: Lista de dicts com {'source': 'ColA', 'destination': 'ColB'}
    """
    data = _load_all()
    data[name] = mapping
    _save_all(data)


def load_layout(name: str) -> Optional[List[Dict]]:
    """
    Carrega um layout pelo nome.

    Returns:
        Lista de mapeamentos ou None se não encontrado.
    """
    data = _load_all()
    return data.get(name)


def delete_layout(name: str) -> bool:
    """
    Remove um layout pelo nome.

    Returns:
        True se removido, False se não existia.
    """
    data = _load_all()
    if name in data:
        del data[name]
        _save_all(data)
        return True
    return False
