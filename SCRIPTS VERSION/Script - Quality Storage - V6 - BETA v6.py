import requests
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pandas as pd
import os
import warnings
from datetime import datetime, timezone, time as dt_time
from datetime import time
import subprocess
import sys
import json

# =======================
# CONFIGURAÇÕES DE CAMINHO INICIAL
# =======================
# Caminho de rede predefinido solicitado pelo usuário
PREDEFINED_NETWORK_PATH = r"\\172.24.80.153\axi-aoi\QUALIDADE\INJET FOLDER LOG\New folder"
# Caminho inicial/default (pode ser o antigo ou um caminho local)
INITIAL_NETWORK_PATH = r"Downloads"


# =======================
# CONFIGURAÇÕES API
# =======================
API_URL_BASE = "https://MAN-prd.jemsms.corp.jabil.org/"
DEFAULT_TIMEOUT = 15
INJET_RESOURCES = [
    "PCB Cleaning INJET IN/OUT",
    "injet OUT BOT",
    "injet OUT TOP",
    "injet IN BOT",
    "injet IN TOP",
    "PCB Cleaning Injet IN",
    "PCB Cleaning Injet OUT"
]

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

_cached_token = None
# =======================
# UTIL / TURNOS
# =======================
def definir_turno(data_hora_obj):
    """
    Retorna o turno (A ou CB) com base no horário informado.
    """
    try:
        data_hora = None
        if isinstance(data_hora_obj, str):
            data_hora_str = data_hora_obj.strip()
            
            try:
                data_hora = datetime.strptime(data_hora_str, "%Y-%m-%d %H:%M:%S %z")
            except ValueError:
                data_hora_str_clean = data_hora_str.replace("T", " ").split(".")[0]
                data_hora_str_clean = data_hora_str_clean.split("+")[0].split("-0")[0] if "Z" not in data_hora_str_clean else data_hora_str_clean.replace("Z", "")
                try:
                    data_hora = datetime.strptime(data_hora_str_clean, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    try:
                        data_hora = datetime.fromisoformat(data_hora_str)
                    except Exception:
                        return "N/A"

        elif isinstance(data_hora_obj, datetime):
            data_hora = data_hora_obj
        else:
            return "N/A"
        
        if data_hora is None:
            return "N/A"

        hora = data_hora.time()

        if time(6, 0) <= hora <= time(15, 48):
            return "A"

        if time(15, 49) <= hora <= time(23, 59):
            return "CB"
        
        if time(0, 0) <= hora < time(6, 0):
            return "CB"
        
        return "N/A"
    
    except Exception:
        return "N/A"

# =======================
# FUNÇÕES DE REDE
# =======================
def ensure_network_connection(network_path):
    """
    Verifica a conexão de rede para o path fornecido.
    """
    if os.path.exists(network_path):
        return True

    messagebox.showwarning(
        "Conexão de Rede",
        f"A pasta de rede não está acessível:\n{network_path}\n\nSerá necessário autenticar..."
    )

    username = simpledialog.askstring("Login de Rede", "Usuário (ex: jabil\\seu_usuario):")
    password = simpledialog.askstring("Senha de Rede", "Senha:", show="*")
    if not username or not password:
        messagebox.showerror("Erro", "Credenciais de rede não fornecidas.")
        return False

    try:
        # Usa o path de rede como alvo para autenticação
        cmd = f'net use "{network_path}" /user:{username} "{password}"'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            messagebox.showinfo("Conectado", f"Conectado com sucesso à pasta de rede:\n{network_path}")
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
def get_token(user, password):
    global _cached_token
    if _cached_token:
        return _cached_token
    url = f"{API_URL_BASE}api-external-api/api/user/adsignin"
    # Usa os parâmetros passados
    form_data = {"name": user, "password": password} 
    try:
        resp = requests.post(url, data=form_data, verify=False, timeout=DEFAULT_TIMEOUT)
        resp.raise_for_status()
        _cached_token = resp.text.strip()
        print("✅ Token obtido")
        return _cached_token
    except Exception as e:
        print("Erro get_token:", e)
        return None


def build_session(user, password):
    token = get_token(user, password)
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
    
    STEPS_TO_FIND = ("SPI BOT", "SPI TOP", "AOI-POS-BOT", "AOI-POS-TOP")
    
    data_list = data.get("Wips", []) if isinstance(data, dict) else []
    
    for wip_entry in data_list:
        # Percorre em ordem reversa para obter a última ocorrência (mais relevante)
        for op in reversed(wip_entry.get("OperationHistories", [])):
            name = op.get("RouteStepName")
            if name in STEPS_TO_FIND and name not in operation_data: 
                operation_data[name] = { 
                    "StartDateTime": op.get("StartDateTime", ""),
                    "Resource": op.get("Resource", "")
                }
    return operation_data


def determine_lado(session, serial):
    """
    Determina se o lado é BOT ou TOP baseado na passagem pelo step AOI-POS-BOT.
    
    Regra:
    - Se a placa passou em AOI-POS-BOT, é considerado TOP.
    - Caso contrário, é considerado BOT.
    """
    op_info = GetOperationHistoryInfo(session, serial)
    
    # Verifica se a chave 'AOI-POS-BOT' existe e tem um valor de tempo associado.
    if "AOI-POS-BOT" in op_info and op_info["AOI-POS-BOT"].get("StartDateTime"):
        return "TOP"
    else:
        return "BOT"


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


def get_area_from_resource(session, resource_name, target_step_name):
    if not session or not resource_name or not target_step_name:
        return ""

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

    for item in data:
        for route in item.get("Routes", []):
            for step in route.get("RouteSteps", []):
                if step.get("RouteStepName") == target_step_name: 
                    areas = step.get("RouteStepManufacturingAreas", [])
                    if areas:
                        return areas[0].get("ManufacturingAreaName", "")
    
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
    
    data_list = data.get("Wips", []) if isinstance(data, dict) else []

    for wip_entry in data_list:
        for op in wip_entry.get("OperationHistories", []):
            resource = op.get("Resource")
            if resource and resource in INJET_RESOURCES:
                injet_entries.append({
                    "Resource": resource,
                    "StartDateTime": op.get("StartDateTime", ""),
                    "RouteStepName": op.get("RouteStepName", ""),
                    "OperationName": op.get("OperationName", ""),
                    "Operator": op.get("Operator", ""),
                    "OperationStatus": op.get("OperationStatus", "")
                })

    for idx, inj in enumerate(injet_entries, 1):
        prefix = f"Injet_{idx}"
        result[f"{prefix}_Resource"] = inj["Resource"]
        result[f"{prefix}_StartDateTime"] = inj["StartDateTime"]
        result[f"{prefix}_RouteStepName"] = inj["RouteStepName"]
        result[f"{prefix}_OperationName"] = inj["OperationName"]
        result[f"{prefix}_Operator"] = inj["Operator"]
        result[f"{prefix}_OperationStatus"] = inj["OperationStatus"]

    return result

# =======================
# EXTRAÇÃO & LOG
# =======================
def extract_info(session, serial):
    info = {}

    wip_info = get_wip_info(session, serial)
    if not wip_info:
        print(f"⚠️ {serial}: não foi possível obter WIP info.")
        return {
            "Serial": serial or "", "Modelo": "", "FERT": "", "Descricao": "",
            "Revision": "", "Versao": "", "Ordem": "", "WipStatus": "",
            "Criado": "", "WipId": None,
            "SPI BOT - Data": "", "SPI BOT - Resource": "", "SPI BOT - Area": "",
            "SPI TOP - Data": "", "SPI TOP - Resource": "", "SPI TOP - Area": "",
            "AOI-POS-BOT - Data": "", "AOI-POS-BOT - Resource": "", "AOI-POS-BOT - Area": "",
            "AOI-POS-TOP - Data": "", "AOI-POS-TOP - Resource": "", "AOI-POS-TOP - Area": "",
        }

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

    op_info = GetOperationHistoryInfo(session, serial)

    spi_bot_resource = op_info.get("SPI BOT", {}).get("Resource", "")
    spi_top_resource = op_info.get("SPI TOP", {}).get("Resource", "")
    
    info["SPI BOT - Data"] = op_info.get("SPI BOT", {}).get("StartDateTime", "")
    info["SPI BOT - Resource"] = spi_bot_resource
    info["SPI TOP - Data"] = op_info.get("SPI TOP", {}).get("StartDateTime", "")
    info["SPI TOP - Resource"] = spi_top_resource

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

    info["SPI BOT - Area"] = get_area_from_resource(session, spi_bot_resource, "SPI BOT")
    info["SPI TOP - Area"] = get_area_from_resource(session, spi_top_resource, "SPI TOP")
    info["AOI-POS-BOT - Area"] = get_area_from_resource(session, aoi_bot_resource, "AOI-POS-BOT")
    info["AOI-POS-TOP - Area"] = get_area_from_resource(session, aoi_top_resource, "AOI-POS-TOP")


    defects = []
    if info.get("WipId"):
        defects = get_defects_by_wip(session, info["WipId"], only_open=False)
    if defects:
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
        info["Defeito(s)"] = ""
        info["CRD(s)"] = ""
        info["Status Defeito(s)"] = ""
        info["FailureLabel(s)"] = ""
        info["Input Defeito(s)"] = ""

    return info


def safe_parse_iso(dt_str):
    if not dt_str or dt_str.strip() == "":
        return ""
    try:
        if isinstance(dt_str, str) and dt_str.endswith("Z"):
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        else:
            dt = datetime.fromisoformat(dt_str) if isinstance(dt_str, str) and "T" in dt_str else datetime.fromtimestamp(float(dt_str))
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return str(dt_str)[:19]


def format_info_line(info, linha_api, lado, serial_oknok, defects, oba_user):
    try:
        agora_dt = datetime.now().astimezone() 
        agora_str = agora_dt.strftime("%Y-%m-%d %H:%M:%S %z")
    except Exception:
        agora_dt = datetime.now()
        agora_str = agora_dt.strftime("%Y-%m-%d %H:%M:%S -0000")

    turno = definir_turno(agora_dt)

    parts = []
    for k, v in info.items():
        if k != "WipId":
            parts.append(f"{k}: {v}")
    
    parts.append(f"Linha: {linha_api}")
    parts.append(f"Lado: {lado}")
    parts.append(f"Status: {serial_oknok}")
    parts.append(f"OBA: {oba_user}") 
    
    if defects:
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
    
    parts.append(f"DataHoraProcessamento: {agora_str}")
    parts.append(f"Turno: {turno}")
    
    return " | ".join(parts)


def append_to_log(session, wip_data, linha_manual_unused, lado, oknok_manual_unused, oba_user, log_file=None, network_path=None):
    if not log_file or not network_path:
        return 0 

    if not ensure_network_connection(network_path):
        return 0

    defects = []
    if wip_data and wip_data.get("WipId"):
        defects = get_defects_by_wip(session, wip_data["WipId"], only_open=False)

    serial_oknok = "OK" 
    forced_nok = any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects)
    if forced_nok:
        serial_oknok = "NOK"

    serial_real = wip_data.get("SerialNumber") if wip_data else None
    
    info = extract_info(session, serial_real) if serial_real else {} 
    
    # Obtém a linha (Manufacturing Area) da API com base no lado determinado
    linha_api = ""
    if lado == "TOP":
        linha_api = info.get("SPI TOP - Area", "")
    else:
        linha_api = info.get("SPI BOT - Area", "")

    # Passa oba_user
    linha_fmt = format_info_line(info, linha_api, lado, serial_oknok, defects, oba_user)
    
    panel_id = get_panel_id_by_serial(session, serial_real)
    panel_id_str = str(panel_id) if panel_id is not None else ""
    linha_fmt += f" | PanelId: {panel_id_str}"
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(linha_fmt + "\n")
        return 1
    except Exception as e:
        try:
            # Em processamento em lote, evitar abrir messagebox na thread worker
            if not threading.current_thread().name.startswith('BatchWorker'):
                 messagebox.showerror("Erro", f"Falha ao escrever log:\n{e}")
            else:
                print(f"Erro ao escrever log de {serial_real} (Lote): {e}")
        except Exception:
            print("Erro ao escrever log:", e)
        return 0


def append_to_log_all_boards(session, panel_data, linha_manual_unused, lado, oknok_manual_unused, oba_user, serial_principal=None, log_file_all=None, network_path=None):
    if not panel_data or not log_file_all or not network_path:
        return 0

    if not ensure_network_connection(network_path):
        return 0

    count = 0
    panel_id_cache = {}

    try:
        with open(log_file_all, "a", encoding="utf-8") as f:
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
                            defects = [{"FailureLabel": "", "DefectName": "", "DefectStatus": "", "Crd": "", "Input": ""}]

                    serial_oknok = "OK"
                    if any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects):
                        serial_oknok = "NOK"

                    info = extract_info(session, serial) if wip else {
                        "Serial": serial, "Modelo": pw.get("MaterialName", ""),
                        "FERT": pw.get("MaterialName", ""), "Descricao": "",
                        "Revision": "", "Versao": "", "Ordem": "",
                        "WipStatus": "", "Criado": ""
                    }

                    linha_api = ""
                    if lado == "TOP":
                        linha_api = info.get("SPI TOP - Area", "")
                    else:
                        linha_api = info.get("SPI BOT - Area", "")

                    if serial in panel_id_cache:
                        panel_id = panel_id_cache[serial]
                    else:
                        panel_id = get_panel_id_by_serial(session, serial)
                        panel_id_cache[serial] = panel_id

                    # Passa oba_user
                    linha_fmt = format_info_line(info, linha_api, lado, serial_oknok, defects, oba_user)
                    
                    panel_id_str = str(panel_id) if panel_id is not None else ""
                    linha_fmt += f" | PanelId: {panel_id_str}"

                    f.write(linha_fmt + "\n")
                    count += 1
        return count
    except Exception as e:
        print("❌ Falha ao escrever log panel:", e)
        return 0

# =======================
# EXPORTAÇÃO EXCEL (Atualizada para nova coluna OBA)
# =======================
def export_from_txt_to_excel(log_files_list, excel_path):
    import tkinter.messagebox as messagebox

    messagebox.showinfo("Exportação Iniciada ⏳", "A exportação do Excel foi iniciada.\n\nPor favor, aguarde...")

    all_parsed_data = []

    for txt_path in log_files_list:
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
    
    # Tratamento para as colunas 'OBA' e 'Turno'
    if "OBA" not in expanded_df.columns:
         expanded_df["OBA"] = ""
    if "Turno" not in expanded_df.columns:
         expanded_df["Turno"] = ""

    # Reordenar colunas
    cols = list(expanded_df.columns)
    cols_to_move = ["OBA", "Turno"]
    
    new_cols = [c for c in cols if c not in cols_to_move]
    
    if "Status" in new_cols:
        try:
            status_idx = new_cols.index("Status")
            for col in reversed(cols_to_move):
                if col in cols:
                    new_cols.insert(status_idx + 1, col)
        except ValueError:
            new_cols.extend(cols_to_move)
    else:
        new_cols.extend(cols_to_move)


    expanded_df = expanded_df[new_cols]

    for col in expanded_df.columns:
        if any(keyword in col.lower() for keyword in ["data", "hora", "time"]):
            expanded_df[col] = expanded_df[col].fillna("")

    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            expanded_df.to_excel(writer, index=False, sheet_name="SPI_Report_All")

        print(f"✅ Excel exportado com sucesso: {excel_path}")
        messagebox.showinfo("Exportação Concluída ✅", f"Arquivo Excel exportado com sucesso!\n\n📂 Caminho:\n{excel_path}")

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
        self.withdraw()
        
        self.api_user_with_domain = None
        self.api_password = None
        self.oba_user = None

        # 💡 Variáveis para controle do lote e progresso
        self.total_serials = 0
        self.processed_count = 0
        self.is_processing = False

        if not self.prompt_for_credentials():
            messagebox.showerror("Erro", "Login cancelado ou credenciais não fornecidas.")
            self.destroy()
            return
            
        self.session = build_session(self.api_user_with_domain, self.api_password)
        if not self.session:
            messagebox.showerror("Erro", "Falha na autenticação do MES. Verifique Usuário/Senha.")
            self.destroy()
            return

        self.deiconify() 
        self.geometry("800x800")
        self.resizable(True, True)

        self.processados = 0
        self.processados_all_boards = 0
        
        self.network_path = INITIAL_NETWORK_PATH
        self.log_file = None
        self.log_file_all = None
        self.xlsx_file = None

        self.create_widgets()
        
        self.set_paths(INITIAL_NETWORK_PATH)

        sys.stdout = self.StdoutRedirector(self.result_text)
        sys.stderr = self.StdoutRedirector(self.result_text, is_error=True)

    def prompt_for_credentials(self):
        """Pede as credenciais ao usuário no início da aplicação."""
        
        try:
            username = simpledialog.askstring(
                "Login MES API", 
                "Usuário (ex: jabil\\seu_usuario):", 
                initialvalue="jabil\\"
            )
            if not username:
                return False
                
            password = simpledialog.askstring(
                "Senha MES API", 
                "Senha:", 
                show="*"
            )
            if not password:
                return False
                
        except Exception:
            return False

        self.api_user_with_domain = username.strip()
        self.api_password = password.strip()
        
        if "\\" in self.api_user_with_domain:
            self.oba_user = self.api_user_with_domain.split("\\")[-1]
        else:
            self.oba_user = self.api_user_with_domain 

        return True

    def set_paths(self, new_path):
        self.network_path = new_path
        self.log_file = os.path.join(new_path, "wip_log.txt")
        self.log_file_all = os.path.join(new_path, "wip_log_all_serials.txt")
        self.xlsx_file = os.path.join(new_path, "wip_log_exportado.xlsx")
        os.makedirs(self.network_path, exist_ok=True)
        self.update_path_label()

    def update_path_label(self):
        if hasattr(self, 'path_label'):
            self.path_label.config(text=f"Pasta de Salvamento: {self.network_path}")

    def set_predefined_network_path(self):
        self.set_paths(PREDEFINED_NETWORK_PATH)
        messagebox.showinfo("Caminho Definido", f"Caminho definido para:\n{self.network_path}")

    def select_custom_network_path(self):
        new_path = filedialog.askdirectory(
            initialdir=self.network_path,
            title="Selecione a Pasta para Salvar os Logs e Excel"
        )
        if new_path:
            self.set_paths(new_path)
            messagebox.showinfo("Caminho Definido", f"Caminho definido para:\n{self.network_path}")

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
        self.start_batch_processing(seriais)
        
    def import_and_process_txt(self):
        if self.is_processing:
            messagebox.showwarning("Atenção", "Um lote já está em processamento. Aguarde a conclusão.")
            return

        file_path = filedialog.askopenfilename(
            defaultextension=".txt",
            filetypes=[("Arquivos de Texto", "*.txt")],
            title="Selecione o arquivo TXT com Seriais (um por linha)"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                seriais = [line.strip() for line in f if line.strip()]
        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Falha ao ler o arquivo:\n{e}")
            return

        if not seriais:
            messagebox.showwarning("Atenção", "O arquivo TXT não contém seriais válidos.")
            return

        self.start_batch_processing(seriais)

    def start_batch_processing(self, serials):
        if self.is_processing:
            messagebox.showwarning("Atenção", "Um lote já está em processamento. Aguarde a conclusão.")
            return

        self.total_serials = len(serials)
        self.processed_count = 0
        self.is_processing = True

        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, f"--- INÍCIO DO PROCESSAMENTO DE LOTE ({self.total_serials} seriais) ---\n")
        self.result_text.configure(state="disabled")
        
        threading.Thread(target=self._process_batch_worker, args=(serials,), daemon=True, name="BatchWorker").start()

    def _update_batch_progress(self, percent, current_serial):
        self.progress["value"] = percent
        self.status_label.config(text=f"Lote em processamento: {self.processed_count}/{self.total_serials} ({percent:.1f}%). Serial atual: {current_serial}")
        self.update()
        
    def _log_serial_result_in_gui(self, serial, resumo, cont_2, success):
        self.result_text.configure(state="normal")
        color_tag = "black" if success else "red"
        status_text = "OK" if success else "FALHA"
        self.result_text.insert(tk.END, f"\n--- {serial} - {status_text} (Boards adicionais: {cont_2}) ---\n", color_tag)
        self.result_text.insert(tk.END, resumo + "\n", color_tag)
        self.result_text.see(tk.END)
        self.result_text.configure(state="disabled")
        
    def _process_batch_worker(self, serials):
        all_ok = True
        
        for serial in serials:
            try:
                result, resumo, cont_2 = self._process_serial_sync(serial)
                if not result:
                    all_ok = False
            except Exception as e:
                result = False
                resumo = f"Erro fatal: {e}"
                cont_2 = 0
                self.after(0, lambda s=serial, r=resumo: self.result_text.insert(tk.END, f"❌ Erro na thread do lote para {s}: {r}\n", "red"))
                all_ok = False
                
            self.processed_count += 1
            progress_percent = (self.processed_count / self.total_serials) * 100
            
            self.after(0, lambda p=progress_percent, s=serial: self._update_batch_progress(p, s))
            self.after(0, lambda s=serial, r=resumo, c2=cont_2, res=result: self._log_serial_result_in_gui(s, r, c2, res))


        final_message = f"✅ Processamento de lote concluído! ({self.total_serials} seriais)"
        if not all_ok:
            final_message = f"⚠️ Processamento de lote concluído, mas com FALHAS. Verifique o log."
            
        self.after(0, lambda: self.status_label.config(text=final_message))
        self.after(0, lambda: self.progress.config(value=100))
        self.is_processing = False
        self.after(0, lambda: self.result_text.insert(tk.END, "\n--- FIM DO PROCESSAMENTO DE LOTE ---\n"))
        self.after(0, lambda: messagebox.showinfo("Processamento Concluído", final_message))


    def _process_serial_sync(self, serial):
        """Função síncrona que processa um único serial (usada por threads worker)."""
        wip = get_wip_info(self.session, serial)
        if not wip:
            return False, f"Serial {serial} não encontrado ou erro na API.", 0

        lado = determine_lado(self.session, serial) 

        cont = append_to_log(self.session, wip, "N/A", lado, None, self.oba_user, self.log_file, self.network_path)
        cont_2 = self.process_panel(self.session, serial, lado)

        if cont == 0:
            return False, f"Serial {serial}: Nenhum dado gravado no log principal.", cont_2

        self.processados += cont
        self.processados_all_boards += cont_2
        
        defects = get_defects_by_wip(self.session, wip.get("WipId"), only_open=False)
        serial_oknok = "OK" 
        if any(((d.get("FailureLabel") or "").strip().upper() == "IMPURITIES") for d in defects):
            serial_oknok = "NOK"

        info = extract_info(self.session, serial)
        linha_api = info.get("SPI TOP - Area", "") if lado == "TOP" else info.get("SPI BOT - Area", "")
        resumo = format_info_line(info, linha_api, lado, serial_oknok, defects, self.oba_user)
        
        return True, resumo, cont_2

    def create_widgets(self):
        pad = 10
        
        # FRAME DE CONFIGURAÇÃO DE PASTA
        frame_path_config = ttk.LabelFrame(self, text="⚙️ Configuração de Pasta", padding=pad)
        frame_path_config.pack(pady=pad, fill="x", padx=pad)
        
        self.path_label = ttk.Label(frame_path_config, text="Pasta de Salvamento: Inicializando...", wraplength=700)
        self.path_label.pack(pady=(0, 10))
        
        frame_path_buttons = ttk.Frame(frame_path_config)
        frame_path_buttons.pack()
        
        ttk.Button(
            frame_path_buttons, 
            text="📁 Definir Pasta INJET (Pré-definida)", 
            command=self.set_predefined_network_path
        ).grid(row=0, column=0, padx=5, pady=5)
        
        ttk.Button(
            frame_path_buttons, 
            text="📂 Selecionar Pasta Customizada", 
            command=self.select_custom_network_path
        ).grid(row=0, column=1, padx=5, pady=5)

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=pad, padx=pad)
        
        ttk.Label(self, text="Serial:").pack(pady=(pad, 2))
        self.serial_entry = ttk.Entry(self, width=50)
        self.serial_entry.pack()
        self.serial_entry.bind("<Return>", lambda e: self.add_serial_to_queue(self.serial_entry.get()))

        frame_botoes = ttk.Frame(self)
        frame_botoes.pack(pady=(5, 10))
        
        ttk.Button(frame_botoes, text="▶ Processar Serial Único", width=25, command=self.on_process).grid(row=0, column=0, padx=5, pady=2)
        ttk.Button(frame_botoes, text="🧩 Processar Fila", width=25, command=self.process_queue).grid(row=0, column=1, padx=5, pady=2)
        
        # 💡 NOVO BOTÃO PARA IMPORTAR TXT
        ttk.Button(
            frame_botoes,
            text="📥 Importar TXT e Processar",
            width=25,
            command=self.import_and_process_txt
        ).grid(row=1, column=0, padx=5, pady=2)
        
        ttk.Button(
            frame_botoes,
            text="📤 Exportar Excel",
            width=25,
            command=lambda: export_from_txt_to_excel([self.log_file, self.log_file_all], self.xlsx_file)
        ).grid(row=1, column=1, padx=5, pady=2)


        ttk.Label(self, text="Fila de Seriais para Processamento:").pack(pady=(10, 2))
        self.serial_queue_listbox = tk.Listbox(self, height=5, width=120)
        self.serial_queue_listbox.pack(pady=(0, 10))

        # RÓTULO DE INFORMAÇÃO AUTOMÁTICA
        info_text = f"Usuário (OBA): {self.oba_user}\n"
        info_text += "Lado (TOP/BOT) e Status (OK/NOK) são determinados AUTOMATICAMENTE pelo histórico do MES."
        ttk.Label(self, text=info_text).pack(pady=(0, 5))

        self.progress = ttk.Progressbar(self, orient="horizontal", length=740, mode="determinate")
        self.progress.pack(pady=(pad, 5))
        self.status_label = ttk.Label(self, text="")
        self.status_label.pack(pady=5)
        self.result_text = tk.Text(self, height=15, width=120, state="disabled", wrap="word")
        self.result_text.pack(pady=5)
        self.result_text.tag_config("red", foreground="red")
        self.result_text.tag_config("black", foreground="black")

    def on_process(self):
        serial = self.serial_entry.get().strip()
        
        if not serial:
            messagebox.showwarning("Atenção", "Digite o número de série.")
            return
        if self.is_processing:
            messagebox.showwarning("Atenção", "Um lote já está em processamento. Aguarde a conclusão ou use a fila.")
            self.progress["value"] = 0
            self.serial_entry.delete(0, tk.END)
            return
            
        # Inicia o processamento de serial único em uma thread
        threading.Thread(target=self._process_single_serial_worker, args=(serial,), daemon=True).start()

    def _process_single_serial_worker(self, serial):
        self.after(0, lambda: self.progress.config(value=10))
        self.after(0, lambda: self.status_label.config(text=f"Buscando e processando serial único: {serial}..."))
        self.after(0, lambda: self.result_text.configure(state="normal"))
        self.after(0, lambda: self.result_text.delete("1.0", tk.END))
        self.after(0, lambda: self.result_text.configure(state="disabled"))
        
        try:
            result, resumo, cont_2 = self._process_serial_sync(serial)
        except Exception as e:
            result = False
            resumo = f"Erro fatal: {e}"
            cont_2 = 0
            
        self.after(0, lambda: self.progress.config(value=100 if result else 0))
        self.after(0, lambda s=serial, r=resumo, c2=cont_2, res=result: self._log_serial_result_in_gui(s, r, c2, res))
        
        final_msg = f"✅ {serial} processado." if result else f"❌ Falha ao processar {serial}. Verifique o log."
        self.after(0, lambda: self.status_label.config(text=final_msg))
        self.after(0, lambda: self.progress.config(value=0))
        self.after(0, lambda: self.serial_entry.delete(0, tk.END))
        if result:
            self.after(0, lambda: messagebox.showinfo("Processamento Concluído", final_msg))

    def process_panel(self, session, serial_principal, lado):
        panel_data = get_panel_info(session, serial_principal)
        if not panel_data:
            print(f"⚠️ Panel não encontrado para {serial_principal}")
            return 0

        # Passa self.oba_user
        return append_to_log_all_boards(
            session, panel_data, "N/A", lado, None, self.oba_user, serial_principal, self.log_file_all, self.network_path
        )

# =======================
# EXECUÇÃO
# =======================
if __name__ == "__main__":
    app = WipApp()
    app.mainloop()