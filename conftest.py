"""Stubs de selenium/tkinter para testar as regras de negócio em qualquer máquina (ex.: CI Linux)."""
import sys
import types
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

for nome in ["selenium", "selenium.webdriver", "selenium.common", "selenium.common.exceptions",
             "selenium.webdriver.chrome", "selenium.webdriver.chrome.options",
             "selenium.webdriver.chrome.service", "selenium.webdriver.common", "selenium.webdriver.common.by"]:
    sys.modules.setdefault(nome, types.ModuleType(nome))
sys.modules["selenium.common.exceptions"].StaleElementReferenceException = Exception
sys.modules["selenium.webdriver.chrome.options"].Options = object
sys.modules["selenium.webdriver.chrome.service"].Service = object


class _By:
    CSS_SELECTOR = "css selector"
    XPATH = "xpath"


sys.modules["selenium.webdriver.common.by"].By = _By
sys.modules["selenium"].webdriver = sys.modules["selenium.webdriver"]

try:
    import tkinter  # noqa: F401
except ImportError:
    tk = types.ModuleType("tkinter")
    tk.Tk = tk.Toplevel = object
    tk.ttk = tk.messagebox = None
    sys.modules["tkinter"] = tk
