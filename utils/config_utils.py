"""配置读写：config.ini + profiles/*.json。"""
import configparser
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE, "config.ini")
PROFILE_DIR = os.path.join(BASE, "profiles")

DEFAULTS = {
    "general": {"output_dir": "output", "last_tool": ""},
    "run": {"workers": "2"},
}


def _cfg():
    cp = configparser.ConfigParser()
    cp.read(CONFIG_PATH, encoding="utf-8")
    return cp


def get(section, key, fallback=""):
    cp = _cfg()
    if cp.has_section(section) and cp.has_option(section, key):
        return cp.get(section, key)
    return DEFAULTS.get(section, {}).get(key, fallback)


def set(section, key, value):
    cp = _cfg()
    if not cp.has_section(section):
        cp.add_section(section)
    cp.set(section, key, str(value))
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        cp.write(f)


def init_config():
    if os.path.exists(CONFIG_PATH):
        return
    cp = configparser.ConfigParser()
    for sec, d in DEFAULTS.items():
        cp[sec] = d
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        cp.write(f)


def list_profiles():
    os.makedirs(PROFILE_DIR, exist_ok=True)
    out = []
    for n in sorted(os.listdir(PROFILE_DIR)):
        if n.endswith(".json"):
            out.append(n[:-5])
    return out


def save_profile(name, data):
    os.makedirs(PROFILE_DIR, exist_ok=True)
    from utils.path_utils import safe_name
    p = os.path.join(PROFILE_DIR, safe_name(name) + ".json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return name


def load_profile(name):
    from utils.path_utils import safe_name
    p = os.path.join(PROFILE_DIR, safe_name(name) + ".json")
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def delete_profile(name):
    from utils.path_utils import safe_name
    p = os.path.join(PROFILE_DIR, safe_name(name) + ".json")
    if os.path.exists(p):
        os.remove(p)
        return True
    return False
