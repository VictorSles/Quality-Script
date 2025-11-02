import requests
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import os
import warnings
from datetime import datetime, timezone
import subprocess
from openpyxl import load_workbook

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# =======================
# CONFIGURAÇÕES
# =======================
API_URL_BASE = "https://MAN-prd.jemsms.corp.jabil.org/"
USER = r"jabil\svchua_jesmapistg"
PASSWORD = "qKzla3oBDA51Ecq=+B2_z"
NETWORK_PATH = r"C:\Users\3808777\OneDrive - Jabil\Imagens"
os.makedirs(NETWORK_PATH, exist_ok=True)
LOG_FILE = os.path.join(NETWORK_PATH, "wip_log.txt")
LOG_FILE_ALL = os.path.join(NETWORK_PATH, "wip_log_all_serials.txt")
XLSX_FILE = os.path.join(NETWORK_PATH, "wip_log_exportado.xlsx")
TXT_FILE = os.path.join(NETWORK_PATH, "resultados_SPI.txt")
LINHAS = ["Rio Jari", "Rio Jutaí", "Rio Negro", "Rio Japurá", "Rio Juruá", "Rio Xingu", "RioTefé"]
LADOS = ["TOP", "BOT"]
OKNOK = ["OK", "NOK"]
_cached_token = None
DEFAULT_TIMEOUT = 15
# =======================
# FUNÇÕES DE REDE
# =======================
def ensure_network_connection():
    if os.path.exists(NETWORK_PATH):
        return True

    messagebox.showwarning(
        "Conexão de Rede",
        f"A pasta de rede não está acessível:\n{NETWORK_PATH}\n\nSerá necessário autenticar..."
    )

    username = simpledialog.askstring("Login de Rede", "Usuário (ex: jabil\\seu_usuario):")
    password = simpledialog.askstring("Senha de Rede", "Senha:", show="*")
    if not username or not password:
        messagebox.showerror("Erro", "Credenciais de rede não fornecidas.")
        return False

    try:
        cmd = f'net use "{NETWORK_PATH}" /user:{username} "{password}"'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            messagebox.showinfo("Conectado", f"Conectado com sucesso à pasta de rede:\n{NETWORK_PATH}")
            return True
        else:
            messagebox.showerror("Erro", f"Falha na autenticação:\n{result.stderr}")
            return False
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao conectar à rede:\n{e}")
        return False

# =======================
# AUTENTICAÇÃO API
# =======================
def get_token():
    global _cached_token
    if _cached_token:
        return _cached_token
    url = f"{API_URL_BASE}api-external-api/api/user/adsignin"
    form_data = {"name": USER, "password": PASSWORD}
    try:
        resp = requests.post(url, data=form_data, verify=False, timeout=DEFAULT_TIMEOUT)
        resp.raise_for_status()
        _cached_token = resp.text.strip()
        print("✅ Token obtido")
        return _cached_token
    except Exception as e:
        print("Erro get_token:", e)
        return None

def build_session():
    token = get_token()
    if not token:
        return None
    cookie_name, cookie_value = ("AuthToken", token)
    if "=" in token:
        cookie_name, cookie_value = token.split("=", 1)
    s = requests.Session()
    s.verify = False
    s.cookies.set(cookie_name.strip(), cookie_value.strip().strip(";"))
    s.headers.update({
        "Cache-Control": "no-cache",
        "User-Agent": "wip-extractor/1.0"
    })
    return s

# =======================
# API WIP INFO
# =======================
def get_wip_info(session, serial, timeout=DEFAULT_TIMEOUT):
    """
    Versão robusta: valida status_code, content-type e parse JSON com mensagem de debug.
    Retorna o primeiro registro (dict) ou None.
    """
    if session is None:
        print("⚠️ get_wip_info: sessão inválida")
        return None
    url = f"{API_URL_BASE}api-external-api/api/Wips/getWipInformationBySerialNumber"
    params = {"SiteName": "MANAUS", "CustomerName": "SAMSUNG", "SerialNumber": serial}
    try:
        r = session.get(url, params=params, timeout=timeout)
    except Exception as e:
        print(f"❌ get_wip_info({serial}): request falhou -> {e}")
        return None

    if r is None:
        print(f"❌ get_wip_info({serial}): resposta nula")
        return None

    if r.status_code != 200:
        print(f"⚠️ get_wip_info({serial}): HTTP {r.status_code} - snippet: { (r.text or '')[:300] }")
        return None

    ct = r.headers.get("Content-Type", "")
    if "application/json" not in ct.lower():
        # tenta ainda parsear caso o content-type esteja faltando, mas avisa
        try:
            data = r.json()
        except Exception:
            print(f"⚠️ get_wip_info({serial}): content-type inesperado: {ct} - snippet: {(r.text or '')[:300]}")
            return None
    try:
        data = r.json()
    except Exception as e:
        print(f"❌ get_wip_info({serial}): erro parse JSON -> {e} - snippet: {(r.text or '')[:300]}")
        return None

    if isinstance(data, list) and data:
        return data[0]
    return None

def get_panel_info(session, serial, timeout=DEFAULT_TIMEOUT):
    """
    Retorna a lista (data) completa retornada pelo endpoint ou None.
    Usado para processar todos os wips do panel.
    """
    if session is None:
        return None
    url = f"{API_URL_BASE}api-external-api/api/Wips/getWipInformationBySerialNumber"
    params = {"SiteName": "MANAUS", "CustomerName": "SAMSUNG", "SerialNumber": serial}
    try:
        r = session.get(url, params=params, timeout=timeout)
        if r.status_code != 200:
            print(f"get_panel_info: status {r.status_code} - { (r.text or '')[:200] }")
            return None
        try:
            data = r.json()
        except Exception as e:
            print(f"get_panel_info: erro parse JSON -> {e} - snippet: {(r.text or '')[:300]}")
            return None
        return data if isinstance(data, list) and data else None
    except Exception as e:
        print("Erro get_panel_info:", e)
        return None

# =======================
# Info By ID / SPI / Resource
# =======================
def GetInfoBySPI(session, serial):
    """
    Retorna dict com possíveis chaves "SPI BOT" e "SPI TOP" contendo StartDateTime e Resource.
    """
    if session is None:
        return {}
    wip_info = get_wip_info(session, serial)
    if not wip_info:
        # já logado no get_wip_info
        return {}
    wip_id = wip_info.get("WipId")
    if not wip_id:
        print(f"⚠️ GetInfoBySPI: WipId não encontrado para serial {serial}")
        return {}

    url = f"{API_URL_BASE}api-external-api/api/Wips/{wip_id}/OperationHistories"
    try:
        response = session.get(url, timeout=DEFAULT_TIMEOUT)
        if response.status_code != 200:
            print(f"⚠️ GetInfoBySPI({serial}): HTTP {response.status_code} - snippet: {(response.text or '')[:300]}")
            return {}
        try:
            data = response.json()
        except Exception as e:
            print(f"❌ GetInfoBySPI({serial}): erro parse JSON -> {e}")
            return {}
    except Exception as e:
        print(f"❌ Erro ao obter OperationHistories: {e}")
        return {}

    spi_data = {}
    if isinstance(data, dict):
        for _, items in data.items():
            for item in items:
                for op in item.get("OperationHistories", []):
                    name = op.get("RouteStepName")
                    if name in ("SPI BOT", "SPI TOP"):
                        spi_data[name] = {
                            "StartDateTime": op.get("StartDateTime", "N/A"),
                            "Resource": op.get("Resource", "N/A")
                        }
    return spi_data
###########################################################################
def get_panel_id_by_serial(session, serial, timeout=DEFAULT_TIMEOUT):
    """
    Retorna o PanelId do serial consultando o endpoint getWipInformationBySerialNumber.
    Retorna None se não encontrado ou erro.
    """
    panel_data = get_panel_info(session, serial, timeout=timeout)
    if not panel_data:
        print(f"⚠️ Serial {serial}: painel não encontrado")
        return None

    # normalmente o JSON tem 'Panel' com 'PanelId'
    first_item = panel_data[0] if isinstance(panel_data, list) and panel_data else None
    if not first_item:
        return None
    panel = first_item.get("Panel")
    if panel:
        panel_id = panel.get("PanelId")
        if panel_id:
            print(f"✅ Serial {serial} -> PanelId: {panel_id}")
            return panel_id
    return None
###########################################################################
def get_spi_bot_manufacturing_area(session, serial):
    """
    Extrai ManufacturingAreaName a partir do Resource da SPI BOT.
    Faz POST para /api/resource/getInfo com {"ResourceName": resource}.
    Retorna string manufacturing area ou None.
    """
    if session is None:
        return None
    spi_info = GetInfoBySPI(session, serial)
    if not spi_info or "SPI BOT" not in spi_info:
        print(f"⚠️ Serial {serial}: SPI BOT não encontrado no GetInfo.")
        return None

    resource_name = spi_info["SPI BOT"].get("Resource")
    if not resource_name:
        print(f"⚠️ Serial {serial}: SPI BOT sem Resource.")
        return None

    url = f"{API_URL_BASE}api-external-api/api/resource/getInfo"
    payload = {"ResourceName": resource_name}
    try:
        r = session.get(url, params=payload, timeout=DEFAULT_TIMEOUT)
        if r.status_code != 200:
            print(f"⚠️ getInfo Resource {resource_name}: HTTP {r.status_code} - snippet: {(r.text or '')[:300]}")
            return None
        try:
            data = r.json()
        except Exception as e:
            print(f"❌ Erro parse JSON Resource getInfo -> {e} - snippet: {(r.text or '')[:300]}")
            return None
    except Exception as e:
        print(f"❌ Erro ao consultar Resource '{resource_name}': {e}")
        return None

    if not isinstance(data, list) or not data:
        print(f"⚠️ Resource getInfo retornou vazio ou formato inesperado para '{resource_name}'")
        return None

    # procura RouteSteps com name "SPI BOT"
    for item in data:
        for route in item.get("Routes", []):
            for step in route.get("RouteSteps", []):
                if step.get("RouteStepName") == "SPI BOT":
                    areas = step.get("RouteStepManufacturingAreas", [])
                    if areas:
                        manufacturing_area = areas[0].get("ManufacturingAreaName")
                        print(f"✅ Serial {serial} | Resource: {resource_name} | ManufacturingAreaName: {manufacturing_area}")
                        return manufacturing_area
    print(f"⚠️ Serial {serial}: nenhum ManufacturingAreaName encontrado")
    return None

# =======================
# GET DEFECTS
# =======================
def get_defects_by_wip(session, wip_id, only_open=False):
    if session is None or not wip_id:
        return []
    endpoint = f"{API_URL_BASE}/api-external-api/api/Wips/ListDefectsByWipId"
    params = {"WipId": wip_id, "OnlyOpenDefects": only_open}
    try:
        response = session.get(endpoint, params=params, timeout=DEFAULT_TIMEOUT)
        if response.status_code != 200:
            print(f"⚠️ Falha ao consultar defeitos (HTTP {response.status_code})")
            return []
        try:
            data = response.json()
        except Exception as e:
            print(f"❌ Erro parse JSON defects -> {e}")
            return []
        defects = []
        if isinstance(data, list):
            for d in data:
                defects.append({
                    "FailureLabel": d.get("FailureLabel", "N/A"),
                    "DefectName": d.get("DefectName", "N/A"),
                    "DefectStatus": d.get("DefectStatus", "N/A"),
                    "Crd": d.get("Crd", "N/A"),
                    "Input": d.get("DefectAnalysisDateTime", None)
                })
        return defects
    except Exception as e:
        print(f"❌ Erro em get_defects_by_wip: {e}")
        return []

# =======================
# EXTRAÇÃO & LOG
# =======================
def extract_info(wip_data):
    sessao = build_session()
    serial = wip_data.get("SerialNumber", "N/A") if wip_data else None
    spi_info = GetInfoBySPI(sessao, serial) if serial else {}

    if not wip_data:
        return {
            "Serial": "N/A", "Modelo": "N/A", "FERT": "N/A", "Descricao": "N/A",
            "Revision": "N/A", "Versao": "N/A", "Ordem": "N/A", "WipStatus": "N/A",
            "Criado": "", "WipId": None,
            "SPI BOT - Data": spi_info.get("SPI BOT", {}).get("StartDateTime", "N/A"),
            "SPI BOT - Resource": spi_info.get("SPI BOT", {}).get("Resource", "N/A"),
            "SPI TOP - Data": spi_info.get("SPI TOP", {}).get("StartDateTime", "N/A"),
            "SPI TOP - Resource": spi_info.get("SPI TOP", {}).get("Resource", "N/A")
        }

    return {
        "Serial": wip_data.get("SerialNumber", "N/A"),
        "Modelo": wip_data.get("MaterialName", "N/A"),
        "FERT": wip_data.get("MaterialName", "N/A"),
        "Descricao": wip_data.get("AssemblyDescription", "N/A"),
        "Revision": wip_data.get("AssemblyRevision", "N/A"),
        "Versao": wip_data.get("AssemblyVersion", "N/A"),
        "Ordem": wip_data.get("PlannedOrderNumber", "N/A"),
        "WipStatus": wip_data.get("WipStatus", "N/A"),
        "Criado": (wip_data.get("WipCreationDate") or "")[:10],
        "WipId": wip_data.get("WipId", None),
        "SPI BOT - Data": spi_info.get("SPI BOT", {}).get("StartDateTime", "N/A"),
        "SPI BOT - Resource": spi_info.get("SPI BOT", {}).get("Resource", "N/A"),
        "SPI TOP - Data": spi_info.get("SPI TOP", {}).get("StartDateTime", "N/A"),
        "SPI TOP - Resource": spi_info.get("SPI TOP", {}).get("Resource", "N/A")
    }

def safe_parse_iso(dt_str):
    if not dt_str or dt_str == "N/A":
        return "N/A"
    try:
        # tenta parse flexível, retorna dd/mm/YYYY HH:MM:SS se possível
        if isinstance(dt_str, str) and dt_str.endswith("Z"):
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        else:
            dt = datetime.fromisoformat(dt_str) if isinstance(dt_str, str) and "T" in dt_str else datetime.fromtimestamp(float(dt_str))
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return str(dt_str)[:19]

def format_info_line(info, linha, lado, oknok=None, defects=None):
    # DataHoraProcessamento no formato Date/Time/Timezone (usa timezone local se disponível)
    try:
        agora = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %z")
    except Exception:
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S -0000")
    parts = [f"{k}: {v}" for k,v in info.items() if k!="WipId"]
    parts.append(f"Linha: {linha}")
    parts.append(f"Lado: {lado}")
    if oknok:
        parts.append(f"Status: {oknok}")
    if defects:
        # usar '; ' como separador consistente
        nomes = "; ".join(d.get("DefectName","N/A") for d in defects)
        crds = "; ".join(d.get("Crd","N/A") for d in defects)
        status = "; ".join(d.get("DefectStatus","N/A") for d in defects)
        failure_labels = "; ".join((d.get("FailureLabel") or "N/A") for d in defects)
        inputs = "; ".join(safe_parse_iso(d.get("Input","N/A")) for d in defects)
        parts.append(f"Defeito(s): {nomes}")
        parts.append(f"CRD(s): {crds}")
        parts.append(f"Status Defeito(s): {status}")
        parts.append(f"FailureLabel(s): {failure_labels}")
        parts.append(f"Input Defeito(s): {inputs}")
    parts.append(f"DataHoraProcessamento: {agora}")
    return " | ".join(parts)

def append_to_log(session, wip_data, linha, lado, oknok=None):
    if not ensure_network_connection():
        return 0
    defects = []
    if wip_data.get("WipId"):
        defects = get_defects_by_wip(session, wip_data["WipId"], only_open=False)

    serial_oknok = oknok
    forced_nok = any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects)
    if forced_nok:
        serial_oknok = "NOK"

    info = extract_info(wip_data)
    serial_real = wip_data.get("SerialNumber") or info.get("Serial") or None
    # passar o serial real para extrair manufacturing area (corrigido)
    linha_spi_bot = get_spi_bot_manufacturing_area(session, serial_real) or linha
    linha_fmt = format_info_line(info, linha_spi_bot, lado, serial_oknok, defects)
    panel_id = get_panel_id_by_serial(session, serial_real)
    linha_fmt += f" | PanelId: {panel_id}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(linha_fmt + "\n")
        return 1
    except Exception as e:
        # A chamada messagebox pode estar em thread; em casos problemáticos, apenas print
        try:
            messagebox.showerror("Erro", f"Falha ao escrever log:\n{e}")
        except Exception:
            print("Erro ao escrever log:", e)
        return 0

def append_to_log_all_boards(session, panel_data, linha, lado, oknok=None):
    if not panel_data or not ensure_network_connection():
        return 0
    count = 0
    try:
        with open(LOG_FILE_ALL, "a", encoding="utf-8") as f:
            for first in panel_data:
                panel_wips = first.get("Panel", {}).get("PanelWips", [])
                for pw in panel_wips:
                    serial = pw.get("SerialNumber")
                    if not serial:
                        continue
                    wip = get_wip_info(session, serial)
                    defects = []
                    if wip and wip.get("WipId"):
                        defects = get_defects_by_wip(session, wip["WipId"], only_open=False)

                    serial_oknok = oknok
                    forced_nok = any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects)
                    if forced_nok:
                        serial_oknok = "NOK"

                    info = extract_info(wip) if wip else {
                        "Serial": serial, "Modelo": pw.get("MaterialName","N/A"),
                        "FERT": pw.get("MaterialName","N/A"), "Descricao":"N/A",
                        "Revision":"N/A","Versao":"N/A","Ordem":"N/A",
                        "WipStatus":"N/A","Criado":""
                    }
                    linha_fmt = format_info_line(info, linha, lado, serial_oknok, defects)
                    f.write(linha_fmt + "\n")
                    count += 1
        return count
    except Exception as e:
        try:
            messagebox.showerror("Erro", f"Falha ao escrever log panel:\n{e}")
        except Exception:
            print("Falha ao escrever log panel:", e)
        return 0

# =======================
# EXPORTAÇÃO EXCEL
# =======================
def format_datetime_with_timezone(value):
    """
    Normaliza uma string/objeto datetime para 'YYYY-MM-DD HH:MM:SS ±HHMM' (mantém - se possível)
    """
    try:
        if not value or str(value).strip() == "":
            return ""
        dt = pd.to_datetime(str(value).strip(), errors="coerce")
        if pd.notnull(dt):
            # localize/format: força formato textual com offset (se tz-aware manter, senão - assume local)
            try:
                tz = dt.tzinfo
                if tz is None:
                    # usa timezone local
                    dt_local = dt.tz_localize(None).astimezone()
                else:
                    dt_local = dt
                return dt_local.strftime("%Y-%m-%d %H:%M:%S %z")
            except Exception:
                return dt.strftime("%Y-%m-%d %H:%M:%S")
        return str(value).strip()
    except Exception:
        return str(value).strip()

def export_from_txt_to_excel(txt_path=LOG_FILE, excel_path=XLSX_FILE):
    """
    Lê o log gerado (separado por ' | ' como "Key: Value"), cria DataFrame,
    expande cada defeito em linhas separadas e exporta para Excel.
    """
    try:
        if not os.path.exists(txt_path):
            messagebox.showerror("Erro", f"Arquivo de log não encontrado:\n{txt_path}")
            return

        # Lê e parseia cada linha
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]

        parsed_data = []
        for line in lines:
            # cada parte formatada como "Key: Value"
            parts = [p.strip() for p in line.split("|") if ":" in p]
            entry = {}
            for p in parts:
                key, _, value = p.partition(":")
                entry[key.strip()] = value.strip()
            parsed_data.append(entry)

        if not parsed_data:
            messagebox.showwarning("Atenção", "Nenhum dado válido encontrado no log.")
            return

        df = pd.DataFrame(parsed_data)

        # normaliza nomes das colunas esperadas (se houver pequenos espaços)
        # Expande defeitos individuais (usar ';' como separador)
        expanded_rows = []
        for _, row in df.iterrows():
            defects = [d.strip() for d in str(row.get("Defeito(s)", "")).split(";") if d.strip()]
            crds = [d.strip() for d in str(row.get("CRD(s)", "")).split(";") if d.strip()]
            statuses = [d.strip() for d in str(row.get("Status Defeito(s)", "")).split(";") if d.strip()]
            failures = [d.strip() for d in str(row.get("FailureLabel(s)", "")).split(";") if d.strip()]
            inputs = [d.strip() for d in str(row.get("Input Defeito(s)", "")).split(";") if d.strip()]

            max_len = max(len(defects), len(crds), len(statuses), len(failures), len(inputs), 1)
            for i in range(max_len):
                new_row = row.to_dict().copy()
                new_row["Defeito(s)"] = defects[i] if i < len(defects) else ""
                new_row["CRD(s)"] = crds[i] if i < len(crds) else ""
                new_row["Status Defeito(s)"] = statuses[i] if i < len(statuses) else ""
                new_row["FailureLabel(s)"] = failures[i] if i < len(failures) else ""
                new_row["Input Defeito(s)"] = inputs[i] if i < len(inputs) else ""
                expanded_rows.append(new_row)

        expanded_df = pd.DataFrame(expanded_rows)

        # padroniza colunas de data/hora
        for col in expanded_df.columns:
            if any(keyword in col.lower() for keyword in ["data", "hora", "time"]):
                expanded_df[col] = expanded_df[col].apply(format_datetime_with_timezone)

        # salva no excel
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            expanded_df.to_excel(writer, index=False, sheet_name="SPI_Report")

        print(f"✅ Excel exportado com sucesso: {excel_path}")
        messagebox.showinfo("✅ Exportação concluída", f"Arquivo Excel salvo com sucesso em:\n{excel_path}")

    except Exception as e:
        print(f"❌ Erro ao exportar: {e}")
        try:
            messagebox.showerror("Erro ao Exportar", str(e))
        except Exception:
            pass

# =======================
# TKINTER APP
# =======================
class WipApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🧠 Jabil MES - WIP Info Extractor")
        self.geometry("650x550")
        self.resizable(False, False)
        self.session = build_session()
        if not self.session:
            messagebox.showerror("Erro", "Falha na autenticação do MES.")
            self.destroy()
            return
        self.processados = 0
        self.processados_all_boards = 0
        self.create_widgets()

    def create_widgets(self):
        pad = 10
        ttk.Label(self, text="Serial:").pack(pady=(pad,2))
        self.serial_entry = ttk.Entry(self, width=45)
        self.serial_entry.pack()
        self.serial_entry.bind("<Return>", lambda e: self.on_process())

        ttk.Label(self, text="Linha:").pack(pady=(pad,2))
        self.linha_var = tk.StringVar(value=LINHAS[0])
        ttk.Combobox(self, textvariable=self.linha_var, state="readonly", values=LINHAS).pack()

        ttk.Label(self, text="Lado da Placa:").pack(pady=(pad,2))
        self.lado_var = tk.StringVar(value=LADOS[0])
        lado_cb = ttk.Combobox(self, textvariable=self.lado_var, state="readonly", values=LADOS)
        lado_cb.pack()
        lado_cb.bind("<<ComboboxSelected>>", self.on_lado_selected)

        ttk.Label(self, text="Status (apenas se TOP):").pack(pady=(pad,2))
        self.oknok_var = tk.StringVar(value=OKNOK[0])
        self.oknok_cb = ttk.Combobox(self, textvariable=self.oknok_var, state="disabled", values=OKNOK)
        self.oknok_cb.pack()

        ttk.Button(self, text="Processar", command=self.on_process).pack(pady=(pad,5))
        ttk.Button(
            self,
            text="📤 Exportar Excel (via Log SPI)",
            command=lambda: export_from_txt_to_excel(LOG_FILE, XLSX_FILE)
        ).pack(pady=5)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=(pad,5))
        self.status_label = ttk.Label(self, text="")
        self.status_label.pack(pady=5)
        self.result_text = tk.Text(self, height=10, width=80)
        self.result_text.pack(pady=5)

    def on_lado_selected(self, event):
        if self.lado_var.get() == "TOP":
            self.oknok_cb.config(state="readonly")
        else:
            self.oknok_cb.config(state="disabled")

    def on_process(self):
        serial = self.serial_entry.get().strip()
        linha = self.linha_var.get().strip()
        lado = self.lado_var.get().strip()
        oknok = self.oknok_var.get() if lado == "TOP" else None
        if not serial:
            messagebox.showwarning("Atenção", "Digite o número de série.")
            return
        threading.Thread(target=self.process_serial, args=(serial, linha, lado, oknok), daemon=True).start()

    def process_serial(self, serial, linha, lado, oknok=None):
        self.progress["value"] = 10
        self.status_label.config(text="Buscando dados...")
        self.update()
        wip = get_wip_info(self.session, serial)
        panel_info = get_panel_info(self.session, serial)
        if not wip:
            # atualizar UI somente no main thread
            self.after(0, lambda: self.status_label.config(text="❌ Serial não encontrado ou erro na API."))
            return
        self.progress["value"] = 50
        cont = append_to_log(self.session, wip, linha, lado, oknok)
        cont_2 = append_to_log_all_boards(self.session, panel_info, linha, lado, oknok)
        self.progress["value"] = 100
        if cont == 0:
            self.after(0, lambda: self.status_label.config(text="⚠️ Nenhum dado gravado."))
        else:
            self.processados += cont
            self.processados_all_boards += cont_2
            defects = get_defects_by_wip(self.session, wip.get("WipId"), only_open=False)
            serial_oknok = oknok
            if any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects):
                serial_oknok = "NOK"
            resumo = format_info_line(extract_info(wip), linha, lado, serial_oknok, defects)
            # atualiza widgets via main thread
            self.after(0, lambda: self._update_ui_after_process(serial, resumo, cont_2))

        self.progress["value"] = 0
        self.serial_entry.delete(0, tk.END)

    def _update_ui_after_process(self, serial, resumo, cont_2):
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, resumo + "\n")
        self.status_label.config(text=f"✅ {serial} processado. Total boards escritos no log 2: {self.processados_all_boards}.")
        messagebox.showinfo("Arquivo Gerado", f"Log individual salvo em:\n{LOG_FILE}\n\nTodos os seriais salvos em:\n{LOG_FILE_ALL}")

# =======================
# EXECUÇÃO
# =======================
if __name__ == "__main__":
    app = WipApp()
    app.mainloop()