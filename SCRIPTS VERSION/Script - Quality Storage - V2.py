import requests
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import os
import warnings
from datetime import datetime
import subprocess
from openpyxl import load_workbook

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# =======================
# CONFIGURAÇÕES
# =======================
API_URL_BASE = "https://MAN-prd.jemsms.corp.jabil.org/"
USER = r"jabil\svchua_jesmapistg"
PASSWORD = "qKzla3oBDA51Ecq=+B2_z"

NETWORK_PATH = r"\\172.24.80.153\axi-aoi\QUALIDADE\INJET FOLDER LOG"
os.makedirs(NETWORK_PATH, exist_ok=True)

LOG_FILE = os.path.join(NETWORK_PATH, "wip_log.txt")
LOG_FILE_ALL = os.path.join(NETWORK_PATH, "wip_log_all_serials.txt")
XLSX_FILE = os.path.join(NETWORK_PATH, "wip_log_exportado.xlsx")

LINHAS = ["Rio Jari", "Rio Jutaí", "Rio Negro", "Rio Japurá", "Rio Juruá", "Rio Xingu", "RioTefé"]
LADOS = ["TOP", "BOT"]
OKNOK = ["OK", "NOK"]

_cached_token = None

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
        resp = requests.post(url, data=form_data, verify=False, timeout=15)
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
    try:
        if "=" in token:
            cookie_name, cookie_value = token.split("=", 1)
        else:
            cookie_name, cookie_value = "AuthToken", token
    except Exception:
        print("Token em formato inesperado:", token)
        return None
    session = requests.Session()
    session.verify = False
    session.cookies.set(cookie_name.strip(), cookie_value.strip(";"))
    session.headers.update({"Cache-Control": "no-cache"})
    return session

# =======================
# API WIP INFO
# =======================
def get_wip_info(session, serial):
    try:
        url = f"{API_URL_BASE}api-external-api/api/Wips/getWipInformationBySerialNumber"
        params = {"SiteName": "MANAUS", "CustomerName": "SAMSUNG", "SerialNumber": serial}
        r = session.get(url, params=params, timeout=15)
        if r.status_code != 200:
            print(f"get_wip_info: status {r.status_code} - {r.text[:200]}")
            return None
        data = r.json()
        return data[0] if isinstance(data, list) and data else None
    except Exception as e:
        print("Erro get_wip_info:", e)
        return None

def get_panel_info(session, serial):
    try:
        url = f"{API_URL_BASE}api-external-api/api/Wips/getWipInformationBySerialNumber"
        params = {"SiteName": "MANAUS", "CustomerName": "SAMSUNG", "SerialNumber": serial}
        r = session.get(url, params=params, timeout=15)
        if r.status_code != 200:
            print(f"get_panel_info: status {r.status_code} - {r.text[:200]}")
            return None
        data = r.json()
        return data if isinstance(data, list) and data else None
    except Exception as e:
        print("Erro get_panel_info:", e)
        return None

# =======================
# GET DEFECTS
# =======================
def get_defects_by_wip(session, wip_id, only_open=False):
    """
    Retorna lista de defeitos. Cada item tem:
    - FailureLabel
    - DefectName
    - DefectStatus
    - Crd
    - Input (FailureCreationDate)
    """
    try:
        endpoint = f"{API_URL_BASE}/api-external-api/api/Wips/ListDefectsByWipId"
        params = {"WipId": wip_id, "OnlyOpenDefects": only_open}
        response = session.get(endpoint, params=params, timeout=15)
        if response.status_code == 200:
            data = response.json()
            defects = []
            if isinstance(data, list):
                for d in data:
                    defects.append({
                        "FailureLabel": d.get("FailureLabel", "N/A"),
                        "DefectName": d.get("DefectName", "N/A"),
                        "DefectStatus": d.get("DefectStatus", "N/A"),
                        "Crd": d.get("Crd", "N/A"),
                        "Input": d.get("FailureCreationDate", "N/A")
                    })
            return defects
        else:
            print(f"⚠️ Falha ao consultar defeitos (HTTP {response.status_code})")
            return []
    except Exception as e:
        print(f"❌ Erro em get_defects_by_wip: {e}")
        return []

# =======================
# EXTRAÇÃO & LOG
# =======================
def extract_info(wip_data):
    if not wip_data:
        return {"Serial":"N/A","Modelo":"N/A","FERT":"N/A","Descricao":"N/A","Revision":"N/A","Versao":"N/A","Ordem":"N/A","WipStatus":"N/A","Criado":""}
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
        "WipId": wip_data.get("WipId", None)
    }

def safe_parse_iso(dt_str):
    if not dt_str or dt_str == "N/A":
        return "N/A"
    try:
        # tentativa simples, suporta Z e offsets
        if dt_str.endswith("Z"):
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        else:
            dt = datetime.fromisoformat(dt_str)
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        # fallback: retorna a string original curta
        return dt_str[:19]

def format_info_line(info, linha, lado, oknok=None, defects=None):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    parts = [f"{k}: {v}" for k,v in info.items() if k!="WipId"]
    parts.append(f"Linha: {linha}")
    parts.append(f"Lado: {lado}")
    if oknok:
        parts.append(f"Status: {oknok}")
    if defects:
        nomes = " - ".join(d.get("DefectName","N/A") for d in defects)
        crds = " - ".join(d.get("Crd","N/A") for d in defects)
        status = " - ".join(d.get("DefectStatus","N/A") for d in defects)
        failure_labels = " - ".join((d.get("FailureLabel") or "N/A") for d in defects)
        inputs = " - ".join(safe_parse_iso(d.get("Input","N/A")) for d in defects)
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

    # usa variável local para evitar mutação externa
    serial_oknok = oknok

    # procura **todos** os FailureLabel por "IMPURITIES"
    forced_nok = any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects)
    if forced_nok:
        serial_oknok = "NOK"

    info = extract_info(wip_data)
    linha_fmt = format_info_line(info, linha, lado, serial_oknok, defects)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(linha_fmt + "\n")
        return 1
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao escrever log:\n{e}")
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

                    # usa variável local por serial para não contaminar próximas iterações
                    serial_oknok = oknok

                    # verifica TODOS os defects para IMPURITIES
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
        messagebox.showerror("Erro", f"Falha ao escrever log panel:\n{e}")
        return 0

# =======================
# EXPORTAÇÃO EXCEL
# =======================
def export_to_excel():
    if not os.path.exists(LOG_FILE):
        messagebox.showerror("Erro", f"Arquivo de log não encontrado:\n{LOG_FILE}")
        return

    try:
        # Ler log principal
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            linhas = [l.strip() for l in f if l.strip()]

        registros = []
        for linha in linhas:
            d = {}
            for part in linha.split(" | "):
                if ": " in part:
                    k, v = part.split(": ", 1)
                    d[k.strip()] = v.strip()
            registros.append(d)

        df = pd.DataFrame(registros)
        if df.empty:
            messagebox.showinfo("Exportar Excel", "Nada para exportar do log principal.")
            return

        cols = ["Serial","Modelo","FERT","Descricao","Revision","Versao","Ordem",
                "WipStatus","Criado","Linha","Lado","Status","Defeito(s)","CRD(s)",
                "Status Defeito(s)","FailureLabel(s)","Input Defeito(s)","DataHoraProcessamento"]
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]

        df.to_excel(XLSX_FILE, index=False, sheet_name="WIP Log")

        # Panel Log
        if os.path.exists(LOG_FILE_ALL):
            with open(LOG_FILE_ALL, "r", encoding="utf-8") as f:
                linhas = [l.strip() for l in f if l.strip()]
            registros_all = []
            for linha in linhas:
                d = {}
                for part in linha.split(" | "):
                    if ": " in part:
                        k, v = part.split(": ", 1)
                        d[k.strip()] = v.strip()
                registros_all.append(d)
            df_all = pd.DataFrame(registros_all)
            for c in cols:
                if c not in df_all.columns:
                    df_all[c] = ""
            df_all = df_all[cols]
            with pd.ExcelWriter(XLSX_FILE, mode="a", engine="openpyxl") as writer:
                df_all.to_excel(writer, index=False, sheet_name="Panel Log")

        messagebox.showinfo("Exportar Excel", f"✅ Exportado com sucesso!\n\nArquivo:\n{XLSX_FILE}")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao exportar para Excel:\n{e}")

# =======================
# TKINTER
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
        pad=10
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
        ttk.Button(self, text="📤 Exportar Excel", command=export_to_excel).pack(pady=5)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=(pad,5))
        self.status_label = ttk.Label(self, text="")
        self.status_label.pack(pady=5)
        self.result_text = tk.Text(self, height=10, width=80)
        self.result_text.pack(pady=5)

    def on_lado_selected(self, event):
        if self.lado_var.get()=="TOP":
            self.oknok_cb.config(state="readonly")
        else:
            self.oknok_cb.config(state="disabled")

    def on_process(self):
        serial=self.serial_entry.get().strip()
        linha=self.linha_var.get().strip()
        lado=self.lado_var.get().strip()
        oknok=self.oknok_var.get() if lado=="TOP" else None
        if not serial:
            messagebox.showwarning("Atenção","Digite o número de série.")
            return
        threading.Thread(target=self.process_serial,args=(serial,linha,lado,oknok),daemon=True).start()

    def process_serial(self, serial, linha, lado, oknok=None):
        self.progress["value"]=10
        self.status_label.config(text="Buscando dados...")
        self.update()
        wip = get_wip_info(self.session, serial)
        panel_info = get_panel_info(self.session, serial)
        if not wip:
            self.status_label.config(text="❌ Serial não encontrado ou erro na API.")
            return
        self.progress["value"]=50
        cont = append_to_log(self.session, wip, linha, lado, oknok)
        cont_2 = append_to_log_all_boards(self.session, panel_info, linha, lado, oknok)
        self.progress["value"]=100
        if cont==0:
            self.status_label.config(text="⚠️ Nenhum dado gravado.")
        else:
            self.processados += cont
            self.processados_all_boards += cont_2
            defects = get_defects_by_wip(self.session, wip.get("WipId"), only_open=False)
            # mostra na tela com o ok/nok forçado aplicado localmente
            serial_oknok = oknok
            if any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects):
                serial_oknok = "NOK"
            resumo = format_info_line(extract_info(wip), linha, lado, serial_oknok, defects)
            self.result_text.delete("1.0",tk.END)
            self.result_text.insert(tk.END,resumo+"\n")
            self.status_label.config(text=f"✅ {serial} processado. Total boards escritos no log 2: {self.processados_all_boards}.")
            messagebox.showinfo("Arquivo Gerado",f"Log individual salvo em:\n{LOG_FILE}\n\nTodos os seriais salvos em:\n{LOG_FILE_ALL}")
        self.progress["value"]=0
        self.serial_entry.delete(0,tk.END)

# =======================
# EXECUÇÃO
# =======================
if __name__=="__main__":
    app = WipApp()
    app.mainloop()
