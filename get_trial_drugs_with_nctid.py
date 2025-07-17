import requests
from typing import List, Tuple, Optional

BASE = "https://clinicaltrials.gov/api/v2"


def _search_nct_id(trial_name: str) -> Optional[str]:
    """
    試験の略称・正式名称（例: "CheckMate 227"）を投げて
    最も関連性が高い候補の NCT ID を返す
    """
    params = {
        "query.titles": trial_name,
        "fields": "NCTId,BriefTitle,protocolSection.armsInterventionsModule.interventions",
        "pageSize": 5,
        "format": "json",
    }

    resp = requests.get(f"{BASE}/studies", params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    studies = data.get("studies", [])
    if not studies:
        return None

    return studies[0]["protocolSection"]["identificationModule"]["nctId"]


def _get_drugs_from_nct(nct_id: str) -> List[str]:
    """
    NCT ID から介入薬の一般名を抽出してリストで返す
    """
    detail_params = {
        # Intervention モジュールだけ取り出せばレスポンスが軽い
        "fields": "protocolSection.armsInterventionsModule.interventions",
        "format": "json",
    }
    det_resp = requests.get(f"{BASE}/studies/{nct_id}", params=detail_params, timeout=30)
    det_resp.raise_for_status()
    det_json = det_resp.json()

    interventions = det_json["protocolSection"]["armsInterventionsModule"]["interventions"]

    # type が "Drug" のものだけ抜き出す
    drugs = [item["name"] for item in interventions if item.get("type") == "DRUG"]

    return drugs


def get_trial_drugs(trial_name: str) -> Tuple[str, List[str]]:
    """
    試験名を渡すと `(NCT ID, [drug, drug, ...])` を返すユーティリティ関数
    """
    nct_id = _search_nct_id(trial_name)
    if not nct_id:
        raise ValueError(f"試験名 '{trial_name}' で NCT ID が見つかりませんでした。")

    drugs = _get_drugs_from_nct(nct_id)
    if not drugs:
        raise RuntimeError(f"NCT {nct_id} で薬剤情報が取得できませんでした。")

    return nct_id, drugs


if __name__ == "__main__":
    trial = "CheckMate 227"
    nct, drug_list = get_trial_drugs(trial)
    print(f"{trial} → NCT ID: {nct}")
    print("使用薬剤:", ", ".join(drug_list))
