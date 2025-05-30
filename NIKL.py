# %% 연령대별 요건 충족 파일 필터링
import os
import json
import pandas as pd
from collections import defaultdict

# ✅ 필터링할 주제 목록
valid_topics = {
    '여행/휴가/휴일/자연휴양지',
    '반려동물/반려용품',
    '우정/성격/MBTI',
    '연애/결혼/가족/관혼상제',
    '생활/주거 환경',
    '회사/학교/학창시절'
}

# ✅ 연령대가 모두 동일한 경우만 반환
def get_common_age(speakers):
    ages = {sp.get("age") for sp in speakers if sp.get("age")}
    return list(ages)[0] if len(ages) == 1 else None

# ✅ JSON 파일 하나에서 조건에 맞는 문서만 추출
def extract_valid_docs_from_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    result = defaultdict(list)

    for doc in data.get("document", []):
        meta = doc.get("metadata", {})
        speakers = meta.get("speaker", [])
        topic = meta.get("topic", "")

        if len(speakers) < 2:
            continue
        if topic not in valid_topics:
            continue
        common_age = get_common_age(speakers)
        if not common_age:
            continue

        result[common_age].append({
            "연령대": common_age,
            "파일명": doc.get("id"),
            "주제": topic,
            "참여자수": len(speakers)
        })

    return result

# ✅ 폴더 내 모든 JSON 파일 처리하고 엑셀 저장
def process_all_jsons_in_directory(folder_path, output_excel):
    total_result = defaultdict(list)

    for filename in os.listdir(folder_path):
        if filename.endswith(".json"):
            file_path = os.path.join(folder_path, filename)
            file_result = extract_valid_docs_from_file(file_path)

            for age, entries in file_result.items():
                total_result[age].extend(entries)

    # 엑셀 저장
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for age, rows in total_result.items():
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=age, index=False)

    print(f"✅ 엑셀 파일 저장 완료: {output_excel}")

process_all_jsons_in_directory(
    folder_path="NIKL_DIALOGUE_2024_v1.0",
    output_excel="연령대별_대화목록.xlsx"
)
# %%
