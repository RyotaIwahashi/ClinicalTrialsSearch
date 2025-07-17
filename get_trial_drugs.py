import requests
from typing import List, Tuple, Optional

BASE = "https://clinicaltrials.gov/api/v2"


def _search_nct_id_and_drugs(trial_name: str) -> Tuple[Optional[str], List[str]]:
    """
    試験名（例: "CheckMate 227"）を検索し、すべてのヒットに登録されている Drug 介入名を重複なく配列で返す
    """
    params = {
        "query.titles": trial_name,  # タイトル全文検索
        "fields": "NCTId,BriefTitle,protocolSection.armsInterventionsModule.interventions",
        "pageSize": 5,
        "format": "json",
    }

    resp = requests.get(f"{BASE}/studies", params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    studies = data.get("studies", [])
    if not studies:
        return None, []

    # すべてのヒットから Drug 介入を収集
    drugs: List[str] = []
    for st in studies:
        interventions = st["protocolSection"].get("armsInterventionsModule", {}).get("interventions", [])
        drugs.extend(iv["name"] for iv in interventions if iv.get("type") == "DRUG")

    # --- 重複削除（順序保持）---
    drugs = list(dict.fromkeys(drugs))

    return drugs


if __name__ == "__main__":
    trial = "CheckMate227"
    drugs = _search_nct_id_and_drugs(trial)
    print("使用薬剤:", ", ".join(drugs))
