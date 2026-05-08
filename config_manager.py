import json
import os
import sys


def _config_path() -> str:
    # Quando rodando como EXE gerado pelo PyInstaller, __file__ aponta para uma
    # pasta temporária interna. Usa o diretório do executável para persistir.
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, 'config.json')


CONFIG_FILE = _config_path()

DEFAULT_CONFIG = {
    "tns_alias":          "",
    "user":               "",
    "password":           "",
    "tns_dir":            "",
    "oracle_client_dir":  "",
    "estab":              "1003",
    "de_para": [
        ["1016619208", "1003"],
        ["1098340331", "1010"],
        ["1098940242", "1011"],
        ["2786503480", "1012"],
        ["2786503498", "1013"],
    ],
}


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(config: dict) -> None:
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
