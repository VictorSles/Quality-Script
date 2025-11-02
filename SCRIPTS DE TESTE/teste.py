import requests as rq
import json
import warnings
import os
import re

sitename = "Manaus"
customer_name = "SAMSUNG"
url_base = "https://man-prd.jemsms.corp.jabil.org/"
USUARIO = R"jabil\3808777"
SENHA = "Ivdscorp@#$2025"

#############################################
        # Pegar S/N
#############################################
serial_number = input("Digite o S/N: ").split()
##### pegar token
def get_token():
    warnings.filterwarnings("ignore", message="Unverified HTTPS request")
    try:
        token = None
        url_token = f"{url_base}api-external-api/api/user/adsignin"
        r = rq.post(url_token, data={"name": USUARIO, "password": SENHA}, verify=False)
        if r.status_code == 200:
            token = r.text.strip()
            statusCode = r.status_code
            return token
        else:
            print(f"Token não obtido: {statusCode}")
    except:
        return 
    print(f"Token não obtido: {statusCode}")

#############################################
            #   Criar Sessão
#############################################
def criar_sessao():
    token = get_token()
    if not token:
        print("Token não existente")
    try:
        if "=" in token:
            nome, valor = token.split("=", 1)
            sessao = rq.Session()
            sessao.verify = False
            sessao.cookies.set(nome, valor)
            sessao.headers.update ()
            sessao.headers.update({"Cache-Control": "no-cache"})
            return sessao
    except:
        return None
#############################################
        # get wipId
#############################################
def wipIdBySerialNumber(sessao, serial):
    url = f"{url_base}api-external-api/api/Wips/GetWipInformationBySerialNumber"
    params_date = {"SiteName":"Manaus", "SerialNumber": serial, "CustomerName":customer_name}
    request = sessao.get(url, params=params_date)
    WipId = None
    isPanelBroker = None
    assemblyId = None
    panelId = None
    InQueueRouteStepRouteName = None
    if request.status_code == 200:
        data_json = request.json()
        if isinstance(data_json, list):
            print(f"""\n============  WIP PROPERIES - '{serial_number}'   ============\n""")
            for i, item in enumerate(data_json, start=0):
                for chave, valor in item.items():
                    if chave == "WipId":
                        WipId = valor
                        print(f"↪{chave} = {valor}")
                    elif chave == "AssemblyId":
                        assemblyId = valor
                    elif chave == "Panel":
                        print(f"""\n============  PANEL PARENTS   ============\n""")
                        for k1, v1 in valor.items():
                            if k1 == "PanelId":
                                panelId = v1
                            if k1 == "PanelWips":
                                for listv1 in v1:
                                    for k11, v11 in listv1.items():
                                        if v11 == WipId:
                                            for k11, v11 in listv1.items():
                                                if k11 == "IsPanelBroken":
                                                    isPanelBroker = v11
                                        else:
                                            print(f"\t\t↪{k11} = {v11}")
                            else:
                                print(f"\t↪{k1} = {v1}")
                    elif chave == "InQueueRouteSteps":
                        valorInquiet = valor
                        print(f"{chave} ↓")
                        for k, v in enumerate(valorInquiet):
                            for k2,v2 in v.items():
                                print(f"\t↪{k2} = {v2}")
                                if k2 == "InQueueRouteStepRouteName":
                                    InQueueRouteStepRouteName = v2
                            #for k, v in enumerate(valorInquiet_2):
                            #    for k2,v2 in v.items():
                            #        print(f"\t↪{k2} = {v2}")
                    else:
                        print(f"↪{chave} = {valor}")
            return WipId, isPanelBroker, assemblyId, panelId, InQueueRouteStepRouteName
    else:
        print(f"S/N não encontrado! - {request.status_code}")
#############################################
        # Get Batch ID
#############################################
def getPanelById(sessao, panelID):
    url = f"{url_base}api-external-api/api/PanelWip/getPanelById"
    params_date = {"Id":{panelID}, "IncludeAbortProcess":True}
    request = sessao.get(url, params=params_date)
    data = request.json()
    batchID = None
    bomId = None
    print(f"""\n============  GET PANEL BY ID   ============\n""")
    for k, v in data.items():
        if k == "PanelType":
            print(f"\t↪{k} ↓")
            for k1,v1 in v.items():
                if k1 == "PanelMappings":
                    print(f"\t↪{k1} ↓")
                    for i, item in enumerate(v1):
                        for k2, v2 in item.items():
                            print(f"\t\t{k2} = {v2}")
                else:
                    print(f"\t↪{k1} = {v1}")
        elif k == "PanelWips":
            print(f"↪{k} ↓")
            for k1, v1 in enumerate(v):
                for k2, v2 in v1.items():
                    if k2 == "BomId":
                        bomId = v2
                    if k2 == "BatchId":
                        batchID = v2
                    print(f"\t↪{k2} = {v2}")
        else:
            print(f"↪{k} = {v}")
    return batchID, bomId
#############################################
        # Assembly ID
#############################################
def getAssemblyId(sessao, AssemblyId):
    url = f"{url_base}api-external-api/api/assembly/{AssemblyId}"
    params_date = {"assemblyId":AssemblyId}
    request = sessao.get(url, params= params_date)
    data = request.json()
    assemblyRevisionId = None
    print(f"""\n============  Assembly ID   ============\n""")
    for k, v in data.items():
        if k == "PanelType":
            for k1, v1 in v.items():
                print(f"\t\t↪{k1} = {v1}")
        else:
            if k == "AssemblyRevisionId":
                assemblyRevisionId = v
            print(f"\t↪{k} = {v}")
    return assemblyRevisionId
#############################################
        # Listar defeitos e tratamento de defeitos
#############################################
def listDefectByWipId(wipid, sessao):
    url = f"{url_base}api-external-api/api/Wips/ListDefectsByWipId"
    params_base = {"WipId": wipid, "OnlyOpenDefects": "False"}
    request = sessao.get(url, params=params_base)
    defects_dict = request.json()
    defect_acumulate = []
    if isinstance(defects_dict, list):
        for i, item in enumerate(defects_dict, start=1):
            for chave, valor in item.items():
                defect_acumulate.append(f"{chave} = {valor}")
        return defect_acumulate
def listDefectFormated(defects):
    lista_defeitos = []
    if isinstance(defects, list):
        for i, item in enumerate(defects, start=1):
            lista_defeitos.append(item)
        return lista_defeitos
#############################################
        # Criar arquivo TXT
##############################################
def createFiletxt(serialnumber, defects):
    try:
        for sn in serialnumber:
            sn = sn.upper()
            with open(f"{sn}.txt", "a", encoding="utf-8") as arquivo:
                for item_defect in defects:
                    arquivo.write(f"{item_defect}\n")
                print("Arquivo criado com sucesso!")
    except:
        print("Erro na geração do TXT")
#############################################
        # Get Route Step
##############################################
def getRouteStepId(sessao, resourcer):
    wipId = wipIdBySerialNumber(criar_sessao(), serial_number)
    wipId_to_get_resourcer = wipId[4]
    url_routeStep = f"{url_base}api-external-api/api/resource/getInfo"
    params_base = {"ResourceName":resourcer}
    request = sessao.get(url_routeStep, params=params_base)
    route_validation = None
    routStepId = None
    if request.status_code == 200:
        infoJson = request.json()
        if isinstance(infoJson, list):
            print(f"""\n============  RESOURCER PROPERIES - '{resourcer}'  ============\n""")
            for itens in infoJson:
                for k,v in itens.items():
                    if v == "":
                        v = "No Information"
                        print(f"↪ {k} = {v}")
                    if k == "ResourceManufacturingAreas":
                        value_rsmnf = v
                        for i3, v2 in enumerate(value_rsmnf):
                            print(f"""\n============  {i3 + 1}º Resourcer Manufacturing Areas   ============\n""")
                            for k3, v3 in v2.items():
                                print(f"↪ {k3} = {v3}")
                    elif k == "Routes":
                        value_rsmnf_2 = v
                        for i4, v3 in enumerate(value_rsmnf_2):
                            print(f"""\n============  {i4 + 1}º Route   ============\n""")
                            for k4, v4 in v3.items():
                                if k4 == "RouteName":
                                    if v4 == wipId_to_get_resourcer:
                                        route_validation = True
                                elif k4 == "RouteSteps" and "route_validation" in locals() and route_validation:
                                    for steps in v4:
                                        route_step_id = steps.get("RouteStepID")
                                        routStepId = route_step_id
                                    del route_validation
                                elif k4 == "RouteSteps":
                                    print(f"\t↪ {k4} ↓")
                                    value_rsmnf_3 = v4
                                    for i5, v4 in enumerate(value_rsmnf_3):
                                        for k5, v5 in v4.items():
                                            if k5 == "RouteStepManufacturingAreas":
                                                print(f"\t↪ {k5} ↓")
                                                value_rsmnf_4 = v5
                                                for i6, v5 in enumerate(value_rsmnf_4):
                                                    for k6, v6 in v5.items():
                                                        print(f"\t\t↪ {k6} = {v6}")
                                            else:
                                                print(f"\t↪ {k5} = {v5}")
                                else:
                                    print(f"    ↪ {k4} = {v4}")
                    else:
                        print(f"↪ {k} = {v}")
    return routStepId
############################################# POST #############################################
def WipMaintanance(sessao, varwipId, batchId, assemblyrevisionId, getRoutStepId, bomID, resourcer):
    url = f"{url_base}/core-application/api/wipmaintenance/saveWip"
    parameters = {"assemblyRevisionId":assemblyrevisionId, "batchId": batchId, "bomId": bomID, "routeStepId": getRoutStepId, "wipId":varwipId}
    request = sessao.post(url, json=parameters)
    if request.status_code in (200,204):
        for sn in serial_number:
            print("+++++++++++++++++++++++++++++++++++++++++++++++++")
            print(f"S/N {sn} movido para {resourcer} com sucesso!")
            print("+++++++++++++++++++++++++++++++++++++++++++++++++")
            print(parameters)
    else:
        print(f"""Movimentação não realizada! {request.status_code}          
              Requets: {parameters}
              """)
#############################################
        # Main calls
#############################################
def choicesToApplicate():
    while True:
        try:
            print("""
            [1] Exportar Informações (.TXT)
            [2] Mover S/N para outra fila
            [3] Fechamento de Defeitos abertos
            """)
            choices = int(input("Escolha uma das opções acima: ").strip())
            if choices >= 4 or choices <= 0 or type(choices) != int:
                print("Escolha Invalida!")
            else:
                return choices
                break
        except Exception as e:
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++")
            print("Erro durante a inserção do valor, tente novamente!")
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++")
#+++++++++++++++++++++++++++++++++++++++++++++++++  Variáveis para WIP Maintanence  +++++++++++++++++++++++++++++++++++++++++++++++++
#
#++++++++++++++++++++++++++++++++++++++++++++  Variáveis para criação de TXT com todas as Informações  ++++++++++++++++++++++++++++++
def CriarArquivo():
    wipId = wipIdBySerialNumber(criar_sessao(), serial_number) #2
    listagem_defeitos = listDefectByWipId(wipId, criar_sessao())
    listagem_defeitos_formated = listDefectFormated(listagem_defeitos)
    createFiletxt(serial_number, listagem_defeitos_formated)
#+++++++++++++++++++++++++++++++++++++++++++++++++  Conficional Final  +++++++++++++++++++++++++++++++++++++++++++++++++
def main():
    def moverSn(resourcer):
        wipId = wipIdBySerialNumber(criar_sessao(), serial_number) #1
        assemblyRevisionId = getAssemblyId(criar_sessao(), wipId[2])
        getrouteStepId = getRouteStepId(criar_sessao(), resourcer)
        ###################
        batchId = getPanelById(criar_sessao(), wipId[3])
        batchId = batchId[0]
        ###################
        bomId = getPanelById(criar_sessao(), wipId[3])
        bomId = bomId[1]
        ###################
        varWipId = wipId[0]
        WipMaintanance(criar_sessao(), varWipId, batchId, assemblyRevisionId, getrouteStepId, bomId, resourcer)
    finalChoice = choicesToApplicate()
    if finalChoice == 2:
        resourcer_var = str(input("Qual o Resourcer que deseja enviar?: ").strip())
        moverSn(resourcer_var)
    else:
        print(f"Blz o primeiro funcionou {finalChoice}")
main()