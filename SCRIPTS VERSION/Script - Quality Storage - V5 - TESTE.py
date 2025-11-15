import requests
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import os
import warnings
from datetime import datetime, timezone, time as dt_time
from datetime import datetime, time
import subprocess
import sys
import json

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
INJET_RESOURCES = [
    "PCB Cleaning INJET IN/OUT",
    "injet OUT BOT",
    "injet OUT TOP",
    "injet IN BOT",
    "injet IN TOP",
    "PCB Cleaning Injet IN",
    "PCB Cleaning Injet OUT"
]

_cached_token = None
DEFAULT_TIMEOUT = 15

# =======================
# UTIL / TURNOS
# =======================
import pandas as _pd

def definir_turno(data_hora_obj):
    """
    Retorna o turno (A ou CB) com base no horário informado.
    Aceita objetos datetime ou strings de data.
    """
    try:
        data_hora = None
        # ✅ Normaliza e tenta converter vários formatos possíveis
        if isinstance(data_hora_obj, str):
            data_hora_str = data_hora_obj.strip()
            
            # Tenta o formato do log (com timezone %z) primeiro
            try:
                # Formato: 2025-11-13 04:12:32 -0400
                data_hora = datetime.strptime(data_hora_str, "%Y-%m-%d %H:%M:%S %z")
            except ValueError:
                # Tenta o formato ISO (da API)
                data_hora_str_clean = data_hora_str.replace("T", " ").split(".")[0]
                data_hora_str_clean = data_hora_str_clean.split("+")[0].split("-0")[0] if "Z" not in data_hora_str_clean else data_hora_str_clean.replace("Z", "")
                try:
                    data_hora = datetime.strptime(data_hora_str_clean, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    try:
                        data_hora = datetime.fromisoformat(data_hora_str) # Última tentativa ISO
                    except Exception:
                        return "N/A" # Falha em todos os parsings de string

        elif isinstance(data_hora_obj, datetime):
            data_hora = data_hora_obj # Aceita objeto datetime diretamente
        else:
            return "N/A"
        
        if data_hora is None:
            return "N/A"

        hora = data_hora.time()

        # 🕕 Turno A: 06:00 → 15:48
        if time(6, 0) <= hora <= time(15, 48):
            return "A"

        # ==========================================================
        # 💡 CORREÇÃO DO GAP (BURACO)
        # O Turno CB agora vai de 15:49 até 05:59 (antes das 06:00)
        # ==========================================================
        
        # 🌙 Turno CB (Parte 1: Tarde/Noite): 15:49 → 23:59
        if time(15, 49) <= hora <= time(23, 59):
            return "CB"
        
        # 🌙 Turno CB (Parte 2: Madrugada): 00:00 → 05:59
        if time(0, 0) <= hora < time(6, 0): # Menor que 06:00
            return "CB"
        
        # ==========================================================

        return "N/A" # Horário não coberto (ex: 15:48:01 até 15:48:59)
    
    except Exception:
        return "N/A"

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

    try:
        data = r.json()
    except Exception as e:
        print(f"❌ get_wip_info({serial}): erro parse JSON -> {e} - snippet: {(r.text or '')[:300]}")
        return None

    if isinstance(data, list) and data:
        return data[0]
    return None


def get_panel_info(session, serial, timeout=DEFAULT_TIMEOUT):
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

# 💡 FUNÇÃO 'GetInfoBySPI' MODIFICADA PARA PEGAR AOI TAMBÉM
def GetOperationHistoryInfo(session, serial):
    if session is None:
        return {}
    wip_info = get_wip_info(session, serial)
    if not wip_info:
        return {}
    wip_id = wip_info.get("WipId")
    if not wip_id:
        print(f"⚠️ GetOperationHistoryInfo: WipId não encontrado para serial {serial}")
        return {}

    url = f"{API_URL_BASE}api-external-api/api/Wips/{wip_id}/OperationHistories"
    try:
        response = session.get(url, timeout=DEFAULT_TIMEOUT)
        if response.status_code != 200:
            print(f"⚠️ GetOperationHistoryInfo({serial}): HTTP {response.status_code} - snippet: {(response.text or '')[:300]}")
            return {}
        try:
            data = response.json()
        except Exception as e:
            print(f"❌ GetOperationHistoryInfo({serial}): erro parse JSON -> {e}")
            return {}
    except Exception as e:
        print(f"❌ Erro ao obter OperationHistories: {e}")
        return {}

    operation_data = {} 
    
    # 💡 Lista de steps que queremos capturar
    STEPS_TO_FIND = ("SPI BOT", "SPI TOP", "AOI-POS-BOT", "AOI-POS-TOP")
    
    if isinstance(data, dict):
        for _, items in data.items():
            for item in items:
                # 💡 Procura pelo histórico mais RECENTE (de trás para frente)
                for op in reversed(item.get("OperationHistories", [])): 
                    name = op.get("RouteStepName")
                    # Só adiciona se ainda não tivermos encontrado (o mais recente)
                    if name in STEPS_TO_FIND and name not in operation_data: 
                        operation_data[name] = { 
                            "StartDateTime": op.get("StartDateTime", ""),
                            "Resource": op.get("Resource", "")
                        }
    return operation_data


def get_panel_id_by_serial(session, serial, timeout=DEFAULT_TIMEOUT):
    panel_data = get_panel_info(session, serial, timeout=timeout)
    if not panel_data:
        print(f"⚠️ Serial {serial}: painel não encontrado")
        return None

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


# 💡 NOVA FUNÇÃO AUXILIAR 
# (Substitui a antiga 'get_spi_bot_manufacturing_area')
def get_area_from_resource(session, resource_name, target_step_name):
    """
    Busca a ManufacturingArea de um Resource, filtrando por um RouteStepName específico.
    """
    if not session or not resource_name or not target_step_name:
        return "" # Retorna "" se os dados de entrada forem inválidos

    url = f"{API_URL_BASE}api-external-api/api/resource/getInfo"
    payload = {"ResourceName": resource_name}
    
    try:
        r = session.get(url, params=payload, timeout=DEFAULT_TIMEOUT)
        if r.status_code != 200:
            print(f"⚠️ get_area_from_resource: HTTP {r.status_code} para {resource_name}")
            return "" 
        data = r.json()
    except Exception as e:
        print(f"❌ Erro em get_area_from_resource: {e}")
        return "" 

    if not isinstance(data, list) or not data:
        return ""

    # Itera pela resposta da API (como o JSON que você forneceu)
    for item in data:
        for route in item.get("Routes", []):
            for step in route.get("RouteSteps", []):
                # Compara com o target_step_name (ex: "AOI-POS-BOT" ou "SPI TOP")
                if step.get("RouteStepName") == target_step_name: 
                    areas = step.get("RouteStepManufacturingAreas", [])
                    if areas:
                        return areas[0].get("ManufacturingAreaName", "")
    
    # Se não encontrar o step_name exato na resposta do recurso
    print(f"⚠️ {resource_name}: não encontrou o step '{target_step_name}' na API /resource/getInfo")
    return "" 

# =======================
# GET DEFECTS
# =======================
def get_defects_by_wip(session, wip_id, only_open=False):
    if session is None or not wip_id:
        return []
    endpoint = f"{API_URL_BASE}api-external-api/api/Wips/ListDefectsByWipId"
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
                # Corrigido: Usar string vazia como default
                defects.append({
                    "FailureLabel": d.get("FailureLabel", ""),
                    "DefectName": d.get("DefectName", ""),
                    "DefectStatus": d.get("DefectStatus", ""),
                    "Crd": d.get("Crd", ""),
                    "Input": d.get("DefectAnalysisDateTime", None)
                })
        return defects
    except Exception as e:
        print(f"❌ Erro em get_defects_by_wip: {e}")
        return []

# =======================
# GET INFO BY INJET RESOURCES
# =======================
def GetInfoByInjetResources(session, serial):
    result = {}
    if session is None:
        return result

    wip_info = get_wip_info(session, serial)
    if not wip_info:
        return result

    wip_id = wip_info.get("WipId")
    if not wip_id:
        return result

    url = f"{API_URL_BASE}api-external-api/api/Wips/{wip_id}/OperationHistories"
    try:
        response = session.get(url, timeout=DEFAULT_TIMEOUT)
        if response.status_code != 200:
            print(f"⚠️ GetInfoByInjetResources({serial}): HTTP {response.status_code}")
            return result
        data = response.json()
    except Exception as e:
        print(f"❌ GetInfoByInjetResources({serial}): erro -> {e}")
        return result

    injet_entries = []
    if isinstance(data, dict):
        for _, items in data.items():
            for item in items:
                for op in item.get("OperationHistories", []):
                    resource = op.get("Resource")
                    if resource and resource in INJET_RESOURCES:
                        # Corrigido: Usar string vazia como default
                        injet_entries.append({
                            "Resource": resource,
                            "StartDateTime": op.get("StartDateTime", ""),
                            "RouteStepName": op.get("RouteStepName", ""),
                            "OperationName": op.get("OperationName", ""),
                        })

    for idx, inj in enumerate(injet_entries, 1):
        prefix = f"Injet_{idx}"
        result[f"{prefix}_Resource"] = inj["Resource"]
        result[f"{prefix}_StartDateTime"] = inj["StartDateTime"]
        result[f"{prefix}_RouteStepName"] = inj["RouteStepName"]
        result[f"{prefix}_OperationName"] = inj["OperationName"]

    return result

# =======================
# EXTRAÇÃO & LOG
# =======================
def extract_info(session, serial):
    info = {}

    wip_info = get_wip_info(session, serial)
    if not wip_info:
        print(f"⚠️ {serial}: não foi possível obter WIP info.")
        # Corrigido: Usar string vazia como default
        return {
            "Serial": serial or "", "Modelo": "", "FERT": "", "Descricao": "",
            "Revision": "", "Versao": "", "Ordem": "", "WipStatus": "",
            "Criado": "", "WipId": None,
            # 💡 Garante que campos AOI existam mesmo em falha
            "SPI BOT - Data": "", "SPI BOT - Resource": "", "SPI BOT - Area": "",
            "SPI TOP - Data": "", "SPI TOP - Resource": "", "SPI TOP - Area": "",
            "AOI-POS-BOT - Data": "", "AOI-POS-BOT - Resource": "", "AOI-POS-BOT - Area": "",
            "AOI-POS-TOP - Data": "", "AOI-POS-TOP - Resource": "", "AOI-POS-TOP - Area": "",
        }

    # Corrigido: Usar string vazia como default
    info["Serial"] = wip_info.get("SerialNumber", serial or "")
    info["Modelo"] = wip_info.get("MaterialName", "")
    info["FERT"] = wip_info.get("MaterialName", "")
    info["Descricao"] = wip_info.get("AssemblyDescription", "")
    info["Revision"] = wip_info.get("AssemblyRevision", "")
    info["Versao"] = wip_info.get("AssemblyVersion", "")
    info["Ordem"] = wip_info.get("PlannedOrderNumber", "")
    info["WipStatus"] = wip_info.get("WipStatus", "")
    info["Criado"] = (wip_info.get("WipCreationDate") or "")[:10]
    info["WipId"] = wip_info.get("WipId", None)

    # 💡 1. Chama a nova função UMA VEZ
    op_info = GetOperationHistoryInfo(session, serial)

    # 💡 2. Pega os dados do SPI (Recurso e Data)
    spi_bot_resource = op_info.get("SPI BOT", {}).get("Resource", "")
    spi_top_resource = op_info.get("SPI TOP", {}).get("Resource", "")
    
    info["SPI BOT - Data"] = op_info.get("SPI BOT", {}).get("StartDateTime", "")
    info["SPI BOT - Resource"] = spi_bot_resource
    info["SPI TOP - Data"] = op_info.get("SPI TOP", {}).get("StartDateTime", "")
    info["SPI TOP - Resource"] = spi_top_resource

    # 💡 3. Pega os dados do AOI (Recurso e Data) (NOVOS CAMPOS)
    aoi_bot_resource = op_info.get("AOI-POS-BOT", {}).get("Resource", "")
    aoi_top_resource = op_info.get("AOI-POS-TOP", {}).get("Resource", "")

    info["AOI-POS-BOT - Data"] = op_info.get("AOI-POS-BOT", {}).get("StartDateTime", "")
    info["AOI-POS-BOT - Resource"] = aoi_bot_resource
    info["AOI-POS-TOP - Data"] = op_info.get("AOI-POS-TOP", {}).get("StartDateTime", "")
    info["AOI-POS-TOP - Resource"] = aoi_top_resource

    injet_info = GetInfoByInjetResources(session, serial)
    if injet_info:
        for k, v in injet_info.items():
            info[k] = v

    # 💡 4. Busca as ÁREAS para os 4 steps usando o helper
    info["SPI BOT - Area"] = get_area_from_resource(session, spi_bot_resource, "SPI BOT")
    info["SPI TOP - Area"] = get_area_from_resource(session, spi_top_resource, "SPI TOP")
    info["AOI-POS-BOT - Area"] = get_area_from_resource(session, aoi_bot_resource, "AOI-POS-BOT")
    info["AOI-POS-TOP - Area"] = get_area_from_resource(session, aoi_top_resource, "AOI-POS-TOP")


    defects = []
    if info.get("WipId"):
        defects = get_defects_by_wip(session, info["WipId"], only_open=False)
    if defects:
        # Corrigido: Usar string vazia como default nas junções
        nomes = "; ".join(d.get("DefectName", "") for d in defects)
        crds = "; ".join(d.get("Crd", "") for d in defects)
        status = "; ".join(d.get("DefectStatus", "") for d in defects)
        failure_labels = "; ".join((d.get("FailureLabel") or "") for d in defects)
        inputs = "; ".join(safe_parse_iso(d.get("Input", "")) for d in defects)
        info["Defeito(s)"] = nomes
        info["CRD(s)"] = crds
        info["Status Defeito(s)"] = status
        info["FailureLabel(s)"] = failure_labels
        info["Input Defeito(s)"] = inputs
    else:
        # Corrigido: Usar string vazia como default para defeitos
        info["Defeito(s)"] = ""
        info["CRD(s)"] = ""
        info["Status Defeito(s)"] = ""
        info["FailureLabel(s)"] = ""
        info["Input Defeito(s)"] = ""

    return info


def safe_parse_iso(dt_str):
    # Corrigido: Retorna "" (string vazia) em vez de "N/A"
    if not dt_str or dt_str == "N/A":
        return ""
    try:
        if isinstance(dt_str, str) and dt_str.endswith("Z"):
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        else:
            dt = datetime.fromisoformat(dt_str) if isinstance(dt_str, str) and "T" in dt_str else datetime.fromtimestamp(float(dt_str))
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return str(dt_str)[:19]


def format_info_line(info, linha, lado, oknok=None, defects=None, link_status=None):
    try:
        # 1. Obter o objeto datetime 'agora'
        agora_dt = datetime.now().astimezone() 
        # 2. Formatar para o log TXT
        agora_str = agora_dt.strftime("%Y-%m-%d %H:%M:%S %z")
    except Exception:
        agora_dt = datetime.now()
        agora_str = agora_dt.strftime("%Y-%m-%d %H:%M:%S -0000")

    # 3. Calcular o Turno A PARTIR DO OBJETO DATETIME
    turno = definir_turno(agora_dt) # Passando o objeto datetime

    # 💡 Os novos campos (AOI) serão incluídos automaticamente
    parts = [f"{k}: {v}" for k, v in info.items() if k != "WipId"]
    
    parts.append(f"Linha: {linha}") # 'linha' aqui é a 'linha_api' (Área do SPI)
    parts.append(f"Lado: {lado}")
    if oknok:
        parts.append(f"Status: {oknok}")
    if defects:
        # Corrigido: Usar string vazia como default nas junções
        nomes = "; ".join(d.get("DefectName", "") for d in defects)
        crds = "; ".join(d.get("Crd", "") for d in defects)
        status = "; ".join(d.get("DefectStatus", "") for d in defects)
        failure_labels = "; ".join((d.get("FailureLabel") or "") for d in defects)
        inputs = "; ".join(safe_parse_iso(d.get("Input", "")) for d in defects)
        parts.extend([
            f"Defeito(s): {nomes}",
            f"CRD(s): {crds}",
            f"Status Defeito(s): {status}",
            f"FailureLabel(s): {failure_labels}",
            f"Input Defeito(s): {inputs}"
        ])
    
    # 4. Adicionar ambos ao TXT
    parts.append(f"DataHoraProcessamento: {agora_str}")
    parts.append(f"Turno: {turno}") # <-- 💡 ADICIONADO AQUI
    
    return " | ".join(parts)


def append_to_log(session, wip_data, linha, lado, oknok=None):
    if not ensure_network_connection():
        return 0
    defects = []
    if wip_data and wip_data.get("WipId"):
        defects = get_defects_by_wip(session, wip_data["WipId"], only_open=False)

    serial_oknok = oknok
    forced_nok = any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects)
    if forced_nok:
        serial_oknok = "NOK"

    serial_real = wip_data.get("SerialNumber") if wip_data else None
    
    # 💡 'info' agora tem TODOS os dados (SPI e AOI)
    info = extract_info(session, serial_real) if serial_real else {} 
    
    # 💡 'linha' passada para format_info_line será a área do SPI
    linha_api = ""
    if lado == "TOP":
        linha_api = info.get("SPI TOP - Area", "") # Pega do dict 'info'
    else:
        linha_api = info.get("SPI BOT - Area", "") # Pega do dict 'info'

    # 💡 A 'linha' da combobox não é mais usada, usamos a 'linha_api'
    linha_fmt = format_info_line(info, linha_api, lado, serial_oknok, defects)
    
    panel_id = get_panel_id_by_serial(session, serial_real)
    # Corrigido: Tratar None em PanelId para garantir que seja string vazia
    panel_id_str = str(panel_id) if panel_id is not None else ""
    linha_fmt += f" | PanelId: {panel_id_str}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(linha_fmt + "\n")
        return 1
    except Exception as e:
        try:
            messagebox.showerror("Erro", f"Falha ao escrever log:\n{e}")
        except Exception:
            print("Erro ao escrever log:", e)
        return 0


def append_to_log_all_boards(session, panel_data, linha, lado, oknok=None, serial_principal=None):
    if not panel_data or not ensure_network_connection():
        return 0

    count = 0
    panel_id_cache = {}

    try:
        with open(LOG_FILE_ALL, "a", encoding="utf-8") as f:
            for item in panel_data:
                panel = item.get("Panel", item)
                panel_wips = panel.get("PanelWips") or []
                if not panel_wips and "SerialNumber" in panel:
                    panel_wips = [panel]

                for pw in panel_wips:
                    serial = pw.get("SerialNumber")
                    if not serial or serial == serial_principal:
                        continue

                    wip = get_wip_info(session, serial)
                    defects = []
                    if wip and wip.get("WipId"):
                        defects = get_defects_by_wip(session, wip["WipId"], only_open=False)
                        if not defects:
                            # Corrigido: Usar string vazia como default
                            defects = [{"FailureLabel": "", "DefectName": "", "DefectStatus": "", "Crd": "", "Input": ""}]

                    serial_oknok = oknok
                    if any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects):
                        serial_oknok = "NOK"

                    # 💡 'info' agora tem TODOS os dados (SPI e AOI)
                    info = extract_info(session, serial) if wip else {
                        "Serial": serial, "Modelo": pw.get("MaterialName", ""),
                        "FERT": pw.get("MaterialName", ""), "Descricao": "",
                        "Revision": "", "Versao": "", "Ordem": "",
                        "WipStatus": "", "Criado": ""
                    }

                    # 💡 'linha' passada para format_info_line será a área do SPI
                    linha_api = ""
                    if lado == "TOP":
                        linha_api = info.get("SPI TOP - Area", "") # Pega do dict 'info'
                    else:
                        linha_api = info.get("SPI BOT - Area", "") # Pega do dict 'info'

                    if serial in panel_id_cache:
                        panel_id = panel_id_cache[serial]
                    else:
                        panel_id = get_panel_id_by_serial(session, serial)
                        panel_id_cache[serial] = panel_id

                    # 💡 Chamada da função 'format_info_line'
                    linha_fmt = format_info_line(info, linha_api, lado, serial_oknok, defects)
                    
                    # Corrigido: Tratar None em PanelId para garantir que seja string vazia
                    panel_id_str = str(panel_id) if panel_id is not None else ""
                    linha_fmt += f" | PanelId: {panel_id_str}"

                    f.write(linha_fmt + "\n")
                    count += 1
        return count
    except Exception as e:
        print("❌ Falha ao escrever log panel:", e)
        return 0

# =======================
# EXPORTAÇÃO EXCEL
# =======================
def format_datetime_with_timezone(value):
    try:
        if not value or str(value).strip() == "":
            return ""
        # Tenta converter o formato com timezone primeiro
        try:
            dt = pd.to_datetime(str(value).strip(), format="%Y-%m-%d %H:%M:%S %z", errors="coerce")
        except Exception:
            dt = pd.to_datetime(str(value).strip(), errors="coerce")
            
        if pd.notnull(dt):
            try:
                # Tenta formatar de volta para o padrão com timezone
                return dt.strftime("%Y-%m-%d %H:%M:%S %z")
            except Exception:
                # Fallback para datas sem timezone
                return dt.strftime("%Y-%m-%d %H:%M:%S")
        return str(value).strip()
    except Exception:
        return str(value).strip()


def export_from_txt_to_excel(log_files=None, excel_path=XLSX_FILE):
    import tkinter.messagebox as messagebox

    messagebox.showinfo("Exportação Iniciada ⏳", "A exportação do Excel foi iniciada.\n\nPor favor, aguarde...")

    if log_files is None:
        log_files = [LOG_FILE, LOG_FILE_ALL]

    all_parsed_data = []

    for txt_path in log_files:
        if not os.path.exists(txt_path):
            print(f"⚠️ Arquivo de log não encontrado: {txt_path}")
            continue
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]

        for line in lines:
            parts = [p.strip() for p in line.split("|") if ":" in p]
            entry = {}
            for p in parts:
                key, _, value = p.partition(":")
                entry[key.strip()] = value.strip()
            all_parsed_data.append(entry)

    if not all_parsed_data:
        print("⚠️ Nenhum dado válido encontrado nos logs.")
        messagebox.showwarning("Nenhum Dado ⚠️", "Nenhum dado válido encontrado nos arquivos de log.")
        return

    df = pd.DataFrame(all_parsed_data)

    # Expande múltiplos defeitos mantendo todas as linhas
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

    # =====================
    # APLICAÇÃO DO TURNO
    # =====================
    # 💡 REMOVIDO: O turno agora é lido diretamente do TXT.
    
    # 💡 ADICIONADO: Se a coluna "Turno" não existir (logs antigos), preenche com "N/A"
    if "Turno" not in expanded_df.columns:
         expanded_df["Turno"] = "N/A"

    # Reposiciona Turno logo após DataHoraProcessamento
    cols = list(expanded_df.columns)
    if "DataHoraProcessamento" in cols and "Turno" in cols:
        try:
            # Tenta mover o Turno para depois da DataHoraProcessamento
            idx = cols.index("DataHoraProcessamento")
            cols.insert(idx + 1, cols.pop(cols.index("Turno")))
            expanded_df = expanded_df[cols]
        except Exception as e:
            print(f"Warn: Não foi possível reordenar coluna Turno. {e}")


    # Padroniza datas/hora
    for col in expanded_df.columns:
        if any(keyword in col.lower() for keyword in ["data", "hora", "time"]):
            # Não formata mais, apenas preenche vazios
            expanded_df[col] = expanded_df[col].fillna("")

    # ADICIONA COLUNA DE RESOURCES ESPECIAIS
    RESOURCE_LIST = INJET_RESOURCES

    def find_special_resources(row):
        resources = row.get("SPI BOT - Resource")
        datetimes = row.get("SPI BOT - Data")
        if pd.isna(resources) or pd.isna(datetimes):
            return "" # Corrigido para ""
        if isinstance(resources, str):
            resources = [r.strip() for r in resources.split(",")]
        if isinstance(datetimes, str):
            datetimes = [d.strip() for d in datetimes.split(",")]
        if len(resources) != len(datetimes):
            return "" # Corrigido para ""
        matched = [dt for res, dt in zip(resources, datetimes) if res in RESOURCE_LIST]
        return "; ".join(matched) if matched else "" # Corrigido para ""

    # Exporta tudo
    try:
        with pd.ExcelWriter(excel_path if (excel_path:=excel_path if 'excel_path' in locals() else XLSX_FILE) else XLSX_FILE, engine="openpyxl") as writer:
            expanded_df.to_excel(writer, index=False, sheet_name="SPI_Report_All")

        print(f"✅ Excel exportado com sucesso: {XLSX_FILE}")
        messagebox.showinfo("Exportação Concluída ✅", f"Arquivo Excel exportado com sucesso!\n\n📂 Caminho:\n{XLSX_FILE}")

    except Exception as e:
        print(f"❌ Erro ao exportar Excel: {e}")
        messagebox.showerror("Erro na Exportação ❌", f"Ocorreu um erro ao exportar o Excel:\n\n{e}")

# =======================
# TKINTER APP
# =======================
class WipApp(tk.Tk):
    class StdoutRedirector:
        def __init__(self, text_widget, is_error=False):
            self.text_widget = text_widget
            self.is_error = is_error

        def write(self, message):
            if message.strip() == "":
                return
            color = "red" if self.is_error else "black"
            try:
                self.text_widget.after(0, lambda: self._write_to_widget(message, color))
            except Exception:
                pass

        def _write_to_widget(self, message, color):
            try:
                self.text_widget.configure(state="normal")
                self.text_widget.insert(tk.END, message + "\n", color)
                self.text_widget.tag_config(color, foreground=color)
                self.text_widget.see(tk.END)
                self.text_widget.configure(state="disabled")
            except Exception:
                pass

        def flush(self):
            pass

    def __init__(self):
        super().__init__()
        self.title("🧠 Jabil MES - WIP Info Extractor")
        self.geometry("800x700")
        self.resizable(True, True)

        self.session = build_session()
        if not self.session:
            messagebox.showerror("Erro", "Falha na autenticação do MES.")
            self.destroy()
            return

        self.processados = 0
        self.processados_all_boards = 0

        self.create_widgets()

        # Redireciona prints (após widget estar criado)
        sys.stdout = self.StdoutRedirector(self.result_text)
        sys.stderr = self.StdoutRedirector(self.result_text, is_error=True)

    def add_serial_to_queue(self, serial):
        serial = serial.strip()
        if serial:
            self.serial_queue_listbox.insert(tk.END, serial)
            self.serial_entry.delete(0, tk.END)
            self.serial_entry.focus()

    def process_queue(self):
        count = self.serial_queue_listbox.size()
        if count == 0:
            messagebox.showinfo("Fila Vazia", "Não há seriais na fila para processar.")
            return

        seriais = [self.serial_queue_listbox.get(i) for i in range(count)]
        self.serial_queue_listbox.delete(0, tk.END)

        linha = self.linha_var.get()
        lado = self.lado_var.get()
        oknok = self.oknok_var.get() if lado == "TOP" else None

        for serial in seriais:
            threading.Thread(target=self.process_serial, args=(serial, linha, lado, oknok), daemon=True).start()

    def create_widgets(self):
        pad = 10
        ttk.Label(self, text="Serial:").pack(pady=(pad, 2))
        self.serial_entry = ttk.Entry(self, width=50)
        self.serial_entry.pack()
        self.serial_entry.bind("<Return>", lambda e: self.add_serial_to_queue(self.serial_entry.get()))

        frame_botoes = ttk.Frame(self)
        frame_botoes.pack(pady=(5, 10))
        ttk.Button(frame_botoes, text="▶ Processar", width=18, command=self.on_process).grid(row=0, column=0, padx=5)
        ttk.Button(frame_botoes, text="🧩 Processar Fila", width=18, command=self.process_queue).grid(row=0, column=1, padx=5)
        ttk.Button(
            frame_botoes,
            text="📤 Exportar Excel",
            width=18,
            command=lambda: export_from_txt_to_excel([LOG_FILE, LOG_FILE_ALL], XLSX_FILE)
        ).grid(row=0, column=2, padx=5)

        ttk.Label(self, text="Fila de Seriais para Processamento:").pack(pady=(10, 2))
        self.serial_queue_listbox = tk.Listbox(self, height=10, width=120)
        self.serial_queue_listbox.pack(pady=(0, 10))

        frame_opcoes = ttk.Frame(self)
        frame_opcoes.pack(pady=(10, 10))
        ttk.Label(frame_opcoes, text="Linha (referência):").grid(row=0, column=0, padx=5, sticky="e")
        self.linha_var = tk.StringVar(value=LINHAS[0])
        ttk.Combobox(frame_opcoes, textvariable=self.linha_var, state="readonly", values=LINHAS, width=25).grid(row=0, column=1, padx=5)
        ttk.Label(frame_opcoes, text="Lado:").grid(row=0, column=2, padx=5, sticky="e")
        self.lado_var = tk.StringVar(value=LADOS[0])
        ttk.Combobox(frame_opcoes, textvariable=self.lado_var, state="readonly", values=LADOS, width=10).grid(row=0, column=3, padx=5)
        ttk.Label(frame_opcoes, text="Status:").grid(row=0, column=4, padx=5, sticky="e")
        self.oknok_var = tk.StringVar(value=OKNOK[0])
        ttk.Combobox(frame_opcoes, textvariable=self.oknok_var, state="readonly", values=OKNOK, width=10).grid(row=0, column=5, padx=5)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=740, mode="determinate")
        self.progress.pack(pady=(pad, 5))
        self.status_label = ttk.Label(self, text="")
        self.status_label.pack(pady=5)
        self.result_text = tk.Text(self, height=20, width=120, state="disabled", wrap="word")
        self.result_text.pack(pady=5)
        self.result_text.tag_config("red", foreground="red")
        self.result_text.tag_config("black", foreground="black")

    def on_process(self):
        serial = self.serial_entry.get().strip()
        linha = self.linha_var.get().strip() # A 'linha' da UI (combobox)
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
        if not wip:
            self.after(0, lambda: self.status_label.config(text="❌ Serial não encontrado ou erro na API."))
            self.progress["value"] = 0 # 💡 Limpa a barra de progresso
            return

        self.progress["value"] = 50
        # 💡 'linha' (da combobox) é passada aqui, mas 'append_to_log' vai usar a área da API
        cont = append_to_log(self.session, wip, linha, lado, oknok)
        cont_2 = self.process_panel(self.session, serial, linha, lado, oknok)
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

            # 💡 'info' agora tem TODOS os dados
            info = extract_info(self.session, serial)
            # 💡 'linha_api' é a área do SPI
            linha_api = info.get("SPI TOP - Area", "") if lado == "TOP" else info.get("SPI BOT - Area", "")
            
            # 💡 A função 'format_info_line' agora inclui o turno automaticamente
            resumo = format_info_line(info, linha_api, lado, serial_oknok, defects)
            
            # 💡 Pega o turno que foi calculado DENTRO do format_info_line
            turno_detectado = "N/A"
            try:
                turno_part = [p for p in resumo.split("|") if "Turno:" in p]
                if turno_part:
                    turno_detectado = turno_part[0].split(":")[1].strip()
            except Exception:
                pass # Mantém "N/A" se falhar

            self.after(0, lambda: self._update_ui_after_process(serial, resumo, cont_2, turno_detectado))

        self.progress["value"] = 0
        self.serial_entry.delete(0, tk.END)

    def _update_ui_after_process(self, serial, resumo, cont_2, turno):
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, resumo + "\n")
        self.result_text.configure(state="disabled")
        self.status_label.config(text=f"✅ {serial} processado. Total boards escritos no log 2: {self.processados_all_boards}.")
        try:
            messagebox.showinfo("Arquivo Gerado", f"Log individual salvo em:\n{LOG_FILE}\n\nTodos os seriais salvos em:\n{LOG_FILE_ALL}\n\nTurno detectado: {turno}")
        except Exception:
            pass

    def process_panel(self, session, serial_principal, linha, lado, oknok=None):
        panel_data = get_panel_info(session, serial_principal)
        if not panel_data:
            print(f"⚠️ Panel não encontrado para {serial_principal}")
            return 0

        count = 0
        panel_id_cache = {}

        for item in panel_data:
            panel = item.get("Panel", item)
            panel_wips = panel.get("PanelWips") or []
            if not panel_wips and "SerialNumber" in panel:
                panel_wips = [panel]

            for pw in panel_wips:
                serial = pw.get("SerialNumber")
                if not serial or serial == serial_principal:
                    continue

                wip = get_wip_info(session, serial)
                if not wip:
                    # Corrigido: Usar string vazia como default
                    info = {
                        "Serial": serial,
                        "Modelo": pw.get("MaterialName", ""),
                        "FERT": pw.get("MaterialName", ""),
                        "Descricao": "",
                        "Revision": "",
                        "Versao": "",
                        "Ordem": "",
                        "WipStatus": "",
                        "Criado": ""
                    }
                else:
                    # 💡 'info' agora tem TODOS os dados
                    info = extract_info(session, serial)

                defects = []
                if wip and wip.get("WipId"):
                    defects = get_defects_by_wip(session, wip["WipId"], only_open=False)
                    if not defects:
                        # Corrigido: Usar string vazia como default
                        defects = [{"FailureLabel": "", "DefectName": "", "DefectStatus": "", "Crd": "", "Input": ""}]

                serial_oknok = oknok
                if any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects):
                    serial_oknok = "NOK"

                # 💡 'linha' passada para format_info_line será a área do SPI
                linha_api = ""
                if lado == "TOP":
                    linha_api = info.get("SPI TOP - Area", "") # Pega do dict 'info'
                else:
                    linha_api = info.get("SPI BOT - Area", "") # Pega do dict 'info'

                if serial in panel_id_cache:
                    panel_id = panel_id_cache[serial]
                else:
                    panel_id = get_panel_id_by_serial(session, serial)
                    panel_id_cache[serial] = panel_id

                # 💡 Chamada da função 'format_info_line'
                linha_fmt = format_info_line(info, linha_api, lado, serial_oknok, defects)
                
                panel_id_str = str(panel_id) if panel_id is not None else ""
                linha_fmt += f" | PanelId: {panel_id_str}"

                try:
                    with open(LOG_FILE_ALL, "a", encoding="utf-8") as f:
                        f.write(linha_fmt + "\n")
                        count += 1
                except Exception as e:
                    print(f"❌ Erro ao gravar log do panel: {e}")
        return count

# =======================
# EXECUÇÃO
# =======================
if __name__ == "__main__":
    app = WipApp()
    app.mainloop()