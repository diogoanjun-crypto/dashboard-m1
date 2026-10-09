# -*- coding: utf-8 -*-
"""
Extrator M1 (base 2931) -> JSON para o dashboard HTML
=====================================================
Baixa o relatório de pacotes da Anjun só da base M1 e gera o arquivo
m1_dados.json (enxuto, sem dados de destinatário) para o dashboard.

Uso:
    python extrator_m1.py --inicio 2026-10-01 --fim 2026-10-08
    python extrator_m1.py                          (pergunta o período)
    python extrator_m1.py --inicio 2026-10-01 --fim 2026-10-08 --mesclar
        -> troca só esses dias dentro do JSON que já existe (estende histórico)

Credenciais (NUNCA no código): variáveis de ambiente ANJUN_TOKEN e ANJUN_COOKIE,
ou um arquivo .env na mesma pasta do script:
    ANJUN_TOKEN=...
    ANJUN_COOKIE=sl-session=...

Requer:
    pip install requests pandas python-calamine urllib3
"""

import argparse
import io
import json
import os
import sys
import time
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

import warnings

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ==================== CONFIGURAÇÃO ====================

BASE_URL = "https://newsupplier.anjunexpress.com"
SITE_ID = 2931
NOME_BASE = "M1"
PRAZO_HORAS = 36
SAIDA_PADRAO = Path(__file__).with_name("m1_dados.json")

MAX_TENTATIVAS = 120
INTERVALO_SEGUNDOS = 15

# Classificação 1º/2º corte (mesma regra do extrator regional)
LIMIAR_GAP_HORAS = 3.5
MIN_PACOTES_SEGUNDO_CORTE = 100

FMT = "%Y-%m-%dT%H:%M:%S"

warnings.filterwarnings("ignore", message="Could not infer format")


def _dt(serie):
    return pd.to_datetime(serie, errors="coerce")

# Como achar cada coluna do Excel: palavras-chave (sem acento, minúsculas),
# em ordem de prioridade, + posição fixa de fallback (A=0, B=1, ...).
# Se algo vier errado, o script imprime todas as colunas: ajuste aqui.
COLUNAS = {
    "rec":    {"chaves": ["tempo de entrada no ponto", "entrada no ponto", "inbound"], "idx": 18},
    "ponto":  {"chaves": ["ponto responsavel"], "idx": 14},
    "ent":    {"chaves": ["tempo de assinatura"], "idx": 20},
    "mot":    {"chaves": ["entregador", "motorista"], "idx": 15},
    "id":     {"chaves": ["numero do waybill", "waybill"], "idx": 0},
    "st":     {"chaves": ["status do pacote"], "idx": 5},
    "cli":    {"chaves": ["nome do cliente"], "idx": 6},
    "tag":    {"chaves": ["tag"], "idx": 8},
    "motivo": {"chaves": ["motivo da ocorrencia"], "idx": 23},
    "cidade": {"chaves": ["cidade do destinatario"], "idx": 32},
    "cep":    {"chaves": ["cep do destinatario"], "idx": 33},
}


# ==================== CREDENCIAIS ====================

def carregar_env():
    arq = Path(__file__).with_name(".env")
    if not arq.exists():
        return
    for linha in arq.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if linha and not linha.startswith("#") and "=" in linha:
            k, v = linha.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def criar_sessao():
    token = os.getenv("ANJUN_TOKEN")
    cookie = os.getenv("ANJUN_COOKIE")
    if not token or not cookie:
        print("❌ Defina ANJUN_TOKEN e ANJUN_COOKIE (variável de ambiente ou arquivo .env).")
        sys.exit(1)

    s = requests.Session()
    retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504],
                    raise_on_status=False)
    s.mount("https://", HTTPAdapter(max_retries=retries, pool_connections=5, pool_maxsize=5))
    s.headers.update({
        "accept": "application/json, text/plain, */*",
        "accept-language": "pt-BR,pt;q=0.9",
        "content-type": "application/json",
        "cookie": cookie,
        "origin": BASE_URL,
        "referer": BASE_URL + "/",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "token": token,
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36",
    })
    return s


# ==================== API ANJUN ====================

def payload_export(inicio, fim):
    return {
        "params": {
            "type": 0, "searchTimeType": 1, "number": "",
            "packageStatus": [2, 3, 4, 9, 29, 10, 11, 12, 30],
            "consigneeId": None, "delivererId": None,
            "siteIdList": [SITE_ID],
            "receiveStateList": [], "reason": "", "problemType": "", "problemTypeSub": "",
            "claimsState": "", "processingResult": "", "businessModel": "",
            "orderType": 0, "waybillNumberList": [],
            "storageStartTime": inicio + " 00:00:00",
            "storageEndTime": fim + " 23:59:59",
        }
    }


def payload_task():
    return {"pageNumber": 1, "pageSize": 50,
            "params": {"startTimeBegin": "", "startTimeEnd": "", "taskTab": "today"}}


def _tarefas(sessao):
    r = sessao.post(f"{BASE_URL}/admin/excelTask/page", json=payload_task(), timeout=(10, 30))
    if r.status_code != 200:
        return []
    return (r.json().get("data") or {}).get("records", [])


def baixar_relatorio(sessao, inicio, fim):
    """Pede o exportAll, espera ficar pronto e devolve os bytes do Excel."""
    t0 = time.time()

    try:
        ja_existentes = {i["fileUrl"] for i in _tarefas(sessao)
                         if i.get("status") == 2 and i.get("fileUrl")}
    except Exception:
        ja_existentes = set()

    print(" 📡 Solicitando geração do relatório...")
    try:
        r = sessao.post(f"{BASE_URL}/admin/package/exportAll",
                        json=payload_export(inicio, fim), timeout=(10, 45))
        if r.status_code != 200:
            print(f" ❌ HTTP {r.status_code} na exportação (token/cookie expirados?)")
            return None
        j = r.json()
        if j.get("code") is not None and j.get("code") != 200:
            print(f" ❌ Erro da API Anjun: {j}")
            return None
    except Exception as e:
        print(f" ❌ Erro de conexão: {e}")
        return None

    url = None
    for n in range(1, MAX_TENTATIVAS + 1):
        time.sleep(INTERVALO_SEGUNDOS)
        dec = int(time.time() - t0)
        print(f"\r   ⏳ Checagem {n:02d}/{MAX_TENTATIVAS} | {dec // 60:02d}m {dec % 60:02d}s",
              end="", flush=True)
        try:
            for i in _tarefas(sessao):
                u = i.get("fileUrl")
                if i.get("status") == 2 and u and u not in ja_existentes:
                    url = u
                    break
        except Exception:
            continue
        if url:
            break
    print()

    if not url:
        print(" 🚨 Timeout: a Anjun não liberou o arquivo em 30 minutos.")
        return None

    print(" ⚡ Baixando arquivo...")
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=(10, 120))
        if r.status_code != 200 or not r.content:
            print(f" ❌ Download falhou (HTTP {r.status_code})")
            return None
        return r.content
    except Exception as e:
        print(f" ❌ Erro no download: {e}")
        return None


# ==================== CONVERSÃO ====================

def _norm(t):
    t = unicodedata.normalize("NFD", str(t))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").lower()
    return " ".join(t.split())


def mapear_colunas(df):
    ncols = {c: _norm(c) for c in df.columns}
    mapa = {}
    for campo, cfg in COLUNAS.items():
        achada = None
        for chave in cfg["chaves"]:
            # 1) nome exatamente igual; 2) nome que contém a chave
            achada = next((c for c, n in ncols.items() if n == chave), None) or \
                     next((c for c, n in ncols.items() if chave in n and len(chave) > 3), None)
            if achada:
                break
        if achada is None and cfg.get("idx") is not None and cfg["idx"] < len(df.columns):
            achada = df.columns[cfg["idx"]]
        mapa[campo] = achada
    return mapa


def classificar_corte(df, col_rec, col_ponto):
    """Mesma lógica do Power Query: maior gap de inbound por ponto + dia."""
    df["_dt"] = _dt(df[col_rec])
    df["_dia"] = df["_dt"].dt.date
    df["corte"] = "1º Corte"
    chaves = [col_ponto, "_dia"] if col_ponto else ["_dia"]

    for _, grupo in df.groupby(chaves):
        g = grupo.dropna(subset=["_dt"]).sort_values("_dt")
        if len(g) <= 1:
            continue
        tempos = g["_dt"].tolist()
        gaps = [(tempos[i + 1] - tempos[i]).total_seconds() / 3600 for i in range(len(tempos) - 1)]
        maior = max(gaps)
        pos = gaps.index(maior) + 1
        if maior >= LIMIAR_GAP_HORAS and len(g) - pos >= MIN_PACOTES_SEGUNDO_CORTE:
            df.loc[g.index[pos:], "corte"] = "2º Corte"
    return df.drop(columns=["_dt", "_dia"])


def _txt(df, col):
    """Texto limpo da coluna (vazio quando não existe ou vier nulo)."""
    if not col:
        return ""
    t = df[col].fillna("").astype(str).str.strip()
    return t.where(~t.str.lower().isin(["nan", "nat", "none"]), "")


def converter(df):
    print("\n 📋 Colunas do Excel:")
    for i, c in enumerate(df.columns):
        print(f"    [{i:02d}] {c}")

    m = mapear_colunas(df)
    print("\n 🧭 Mapeamento usado:")
    for campo, col in m.items():
        print(f"    {campo:<6} -> {col if col else '(não encontrada)'}")

    if not m["rec"]:
        print(" ❌ Coluna de recebimento não encontrada. Ajuste COLUNAS['rec'].")
        return None
    if not m["ent"]:
        print(" ⚠️  Coluna de assinatura não encontrada: todos os pacotes ficarão como pendentes.")

    amostra = df[m["rec"]].dropna().astype(str).head(3).tolist()
    print(f"\n 🕒 Exemplo de valores em '{m['rec']}': {amostra}")
    df = classificar_corte(df.copy(), m["rec"], m["ponto"])

    if m["id"]:
        df = df.drop_duplicates(subset=[m["id"]])

    rec = _dt(df[m["rec"]])
    ent = _dt(df[m["ent"]]) if m["ent"] else pd.Series(pd.NaT, index=df.index)

    out = pd.DataFrame({
        "id": df[m["id"]].astype(str) if m["id"] else df.index.astype(str),
        "rec": rec.dt.strftime(FMT),
        "ent": ent.dt.strftime(FMT),
        "mot": _txt(df, m["mot"]),
        "st": _txt(df, m["st"]),
        "cli": _txt(df, m["cli"]),
        "tag": _txt(df, m["tag"]),
        "ponto": _txt(df, m["ponto"]),
        "motivo": _txt(df, m["motivo"]),
        "cidade": _txt(df, m["cidade"]),
        "cep": _txt(df, m["cep"]),
        "corte": df["corte"],
    })
    out = out[rec.notna()].astype(object)
    out = out.where(out.notna(), None)
    return list(out.columns), out.values.tolist(), m


# ==================== SAÍDA ====================

def gravar(saida, cols, rows, mapa, inicio, fim, mesclar):
    i_rec = cols.index("rec")

    if mesclar and saida.exists():
        antigo = json.loads(saida.read_text(encoding="utf-8"))
        if antigo.get("cols") == cols:
            mantidas = [r for r in antigo["rows"] if not (inicio <= (r[i_rec] or "")[:10] <= fim)]
            print(f" 🔀 Mesclando: {len(mantidas):,} linhas antigas mantidas + {len(rows):,} novas")
            rows = mantidas + rows
        else:
            print(" ⚠️  Colunas do JSON antigo diferem; gravando só o período novo.")

    rows.sort(key=lambda r: r[i_rec] or "")
    dias = [r[i_rec][:10] for r in rows if r[i_rec]]
    meta = {
        "base": NOME_BASE,
        "site_id": SITE_ID,
        "gerado_em": datetime.now().strftime(FMT),
        "periodo": [min(dias), max(dias)] if dias else [inicio, fim],
        "prazo_horas": PRAZO_HORAS,
        "total": len(rows),
        "colunas_mapeadas": {k: v for k, v in mapa.items()},
    }
    saida.write_text(
        json.dumps({"meta": meta, "cols": cols, "rows": rows}, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    print(f" 🟢 Gerado: {saida}  ({len(rows):,} linhas, {saida.stat().st_size / 1024:.0f} KB)")


# ==================== MAIN ====================

def data_valida(txt):
    try:
        return datetime.strptime(txt, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        raise argparse.ArgumentTypeError(f"Data inválida: {txt} (use AAAA-MM-DD)")


def perguntar_data(rotulo, padrao):
    while True:
        txt = input(f" {rotulo} [AAAA-MM-DD, Enter = {padrao}]: ").strip() or padrao
        try:
            return data_valida(txt)
        except argparse.ArgumentTypeError as e:
            print(f" ❌ {e}")


def main():
    carregar_env()
    ap = argparse.ArgumentParser(description=f"Extrator {NOME_BASE} -> JSON do dashboard")
    ap.add_argument("--inicio", type=data_valida)
    ap.add_argument("--fim", type=data_valida)
    ap.add_argument("--mesclar", action="store_true",
                    help="troca só o período informado dentro do JSON existente")
    ap.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = ap.parse_args()

    hoje = datetime.now().strftime("%Y-%m-%d")
    padrao_ini = (datetime.now() - timedelta(days=6)).strftime("%Y-%m-%d")
    inicio = args.inicio or perguntar_data("Data início", padrao_ini)
    fim = args.fim or perguntar_data("Data fim   ", hoje)

    if inicio > fim:
        print("❌ Data início maior que a data fim.")
        sys.exit(1)
    if fim > hoje:
        print("❌ Data fim não pode ser futura.")
        sys.exit(1)

    print(f"\n 🚀 {NOME_BASE} (site {SITE_ID}) | {inicio} até {fim}\n")
    sessao = criar_sessao()

    conteudo = baixar_relatorio(sessao, inicio, fim)
    if not conteudo:
        sys.exit(1)

    print(" 📖 Lendo Excel...")
    try:
        df = pd.read_excel(io.BytesIO(conteudo), engine="calamine", dtype=str, header=0)
    except ImportError:
        df = pd.read_excel(io.BytesIO(conteudo), engine="openpyxl", dtype=str, header=0)

    if df.empty:
        print(" ⚠️ O relatório veio com 0 linhas para esse período.")
        sys.exit(1)

    resultado = converter(df)
    if not resultado:
        sys.exit(1)
    cols, rows, mapa = resultado

    gravar(args.saida, cols, rows, mapa, inicio, fim, args.mesclar)


if __name__ == "__main__":
    main()
