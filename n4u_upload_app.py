"""엔포유 소분 작업 및 감모 관리 업로드 프로그램.

실행: streamlit run n4u_upload_app.py
"""

from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
import re
from typing import Any, BinaryIO

import openpyxl
import pandas as pd
import streamlit as st


DISPLAY_COLUMNS = [
    "거래일자", "전표번호", "투입품목코드", "투입품목명", "투입수량",
    "투입중량(kg)", "투입원가(합계)", "산출품목코드", "산출품목명",
    "제조수량", "제조중량(kg)", "납품원가(합계)", "감모수량(kg)",
    "감모금액(원)", "감모비율(%)", "작업비비율(%)", "소분작업비용(합계)",
]


def normalize_code(value: Any) -> str:
    """엑셀의 숫자/문자 제품코드를 같은 문자열로 맞춘다."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text[:-2] if re.fullmatch(r"\d+\.0", text) else text


def to_number(value: Any, default: float = 0.0) -> float:
    if value is None or str(value).strip() == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^0-9.+-]", "", str(value).replace(",", ""))
    try:
        return float(cleaned)
    except ValueError:
        return default


def find_columns(rows: list[tuple[Any, ...]]) -> tuple[int, dict[str, int]]:
    """소분리스트 첫 15행에서 보고서 헤더와 필요한 열을 찾는다."""
    patterns = {
        "date": ["거래년월일", "거래일자", "일자"],
        "doc": ["전표번호", "전표no", "전표"],
        "warehouse": ["창고명", "창고"],
        "qty": ["수량"],
        "name": ["상품명", "품명", "내역", "품목명"],
        "code": ["품목코드", "상품코드", "코드"],
    }
    for row_index, row in enumerate(rows[:15]):
        found: dict[str, int] = {}
        for col_index, cell in enumerate(row):
            header = str(cell or "").replace(" ", "").lower()
            for key, words in patterns.items():
                if key not in found and any(word in header for word in words):
                    found[key] = col_index
            # 원가 계산에는 부가세까지 포함된 합계금액을 사용한다. 합계금액이
            # 없는 옛 양식에서는 공급가액, 판매금액 순으로 대체한다.
            if header == "합계금액":
                found["price"] = col_index
            elif header == "공급가액" and "price" not in found:
                found["price"] = col_index
            elif header == "판매금액" and "price" not in found:
                found["price"] = col_index
        if "warehouse" in found and "price" in found:
            return row_index, found
    raise ValueError("소분리스트에서 필수 열인 '창고명'과 '판매금액/금액'을 찾지 못했습니다.")


def extract_date(value: Any) -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    match = re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}", str(value or ""))
    return match.group().replace(".", "-").replace("/", "-") if match else ""


def first_match(row: tuple[Any, ...], pattern: str) -> str:
    for cell in row:
        match = re.search(pattern, str(cell or ""))
        if match:
            return match.group()
    return ""


def read_product_weights(file: BinaryIO) -> dict[str, float]:
    workbook = openpyxl.load_workbook(file, data_only=True, read_only=True)
    sheet = workbook.active
    weights: dict[str, float] = {}
    for row in sheet.iter_rows(values_only=True):
        if len(row) <= 26:
            continue
        code = normalize_code(row[1])  # B: 제품코드
        weight = to_number(row[26], default=-1)  # AA: 중량
        if code and weight >= 0:
            weights[code] = weight
    return weights


def process_files(list_file: BinaryIO, product_file: BinaryIO) -> list[dict[str, Any]]:
    weights = read_product_weights(product_file)
    workbook = openpyxl.load_workbook(list_file, data_only=True, read_only=True)
    rows = list(workbook.active.iter_rows(values_only=True))
    header_index, columns = find_columns(rows)

    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    current_doc, current_date = "-", "-"
    for row in rows[header_index + 1 :]:
        if not row or len(row) <= max(columns["warehouse"], columns["price"]):
            continue
        row_text = " ".join(str(cell) for cell in row if cell is not None)
        if any(word in row_text for word in ("합계", "합 계", "소계", "소 계")):
            continue
        warehouse = str(row[columns["warehouse"]] or "").strip()
        raw_price = row[columns["price"]]
        if not warehouse or raw_price is None:
            continue

        doc = first_match((row[columns["doc"]],), r"\d{6,}") if "doc" in columns else ""
        doc = doc or first_match(row, r"\d{6,}") or current_doc
        current_doc = doc
        transaction_date = extract_date(row[columns["date"]]) if "date" in columns else ""
        transaction_date = transaction_date or first_match(row, r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}")
        transaction_date = transaction_date.replace(".", "-").replace("/", "-") if transaction_date else current_date
        current_date = transaction_date
        if transaction_date == "-":
            continue

        qty = to_number(row[columns["qty"]], 1.0) if "qty" in columns else 1.0
        item_code = normalize_code(row[columns["code"]]) if "code" in columns else ""
        item_name = str(row[columns["name"]] or "").strip() if "name" in columns else ""
        unit_weight = weights.get(item_code, 0.0)
        detail = {
            "code": item_code or "-", "name": item_name or "-", "qty": qty,
            "price": to_number(raw_price), "unit_weight": unit_weight,
            "weight_kg": qty * unit_weight / 1000,
        }
        record = grouped.setdefault((doc, transaction_date), {
            "inp": [], "out": [],
        })
        if "소분" in warehouse:
            record["inp"].append(detail)
        elif "대연점" in warehouse:
            record["out"].append(detail)

    result: list[dict[str, Any]] = []
    for (doc, transaction_date), record in sorted(grouped.items(), key=lambda item: (item[0][1], item[0][0])):
        def total(side: str, key: str) -> float:
            return sum(float(item[key]) for item in record[side])

        inp_cost, out_cost = total("inp", "price"), total("out", "price")
        inp_weight, out_weight = total("inp", "weight_kg"), total("out", "weight_kg")
        loss_weight = inp_weight + out_weight
        work_cost = inp_cost + out_cost
        unit_price = abs(inp_cost) / abs(inp_weight) if inp_weight else 0.0
        result.append({
            "date": transaction_date, "doc_no": doc,
            "inp_code": ", ".join(sorted({item["code"] for item in record["inp"] if item["code"] != "-"})) or "-",
            "inp_name": ", ".join(sorted({item["name"] for item in record["inp"] if item["name"] != "-"})) or "-",
            "inp_qty": total("inp", "qty"), "inp_cost": inp_cost, "inp_wt_kg": inp_weight,
            "out_code": ", ".join(sorted({item["code"] for item in record["out"] if item["code"] != "-"})) or "-",
            "out_name": ", ".join(sorted({item["name"] for item in record["out"] if item["name"] != "-"})) or "-",
            "out_qty": total("out", "qty"), "out_cost": out_cost, "out_wt_kg": out_weight,
            "loss_wt_kg": loss_weight, "loss_cost": loss_weight * unit_price,
            "loss_ratio": loss_weight / abs(inp_weight) * 100 if inp_weight else 0.0,
            "ratio": work_cost / -inp_cost * 100 if inp_cost else 0.0,
            "work_cost": work_cost, "inp_details": record["inp"], "out_details": record["out"],
        })
    return result


def display_dataframe(records: list[dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for item in records:
        rows.append({
            "거래일자": item["date"], "전표번호": item["doc_no"], "투입품목코드": item["inp_code"],
            "투입품목명": item["inp_name"], "투입수량": item["inp_qty"], "투입중량(kg)": item["inp_wt_kg"],
            "투입원가(합계)": item["inp_cost"], "산출품목코드": item["out_code"], "산출품목명": item["out_name"],
            "제조수량": item["out_qty"], "제조중량(kg)": item["out_wt_kg"], "납품원가(합계)": item["out_cost"],
            "감모수량(kg)": item["loss_wt_kg"], "감모금액(원)": item["loss_cost"],
            "감모비율(%)": item["loss_ratio"], "작업비비율(%)": item["ratio"], "소분작업비용(합계)": item["work_cost"],
        })
    return pd.DataFrame(rows, columns=DISPLAY_COLUMNS)


def to_excel(dataframe: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="소분작업분석")
    return output.getvalue()


st.set_page_config(page_title="엔포유 소분 작업 관리", page_icon=":material/scale:", layout="wide")
st.title("엔포유 소분 작업 및 감모 관리")
st.caption("소분리스트와 제품조회 파일을 업로드하면 전표별 투입·산출·감모 현황을 계산합니다.")

for key, default in {"records": None, "file_signature": None}.items():
    st.session_state.setdefault(key, default)

with st.container(border=True):
    st.subheader("파일 업로드", divider="gray")
    upload_list, upload_product = st.columns(2)
    with upload_list:
        list_file = st.file_uploader("1. 엔포유소분리스트", type=["xlsx"], key="list_file")
    with upload_product:
        product_file = st.file_uploader("2. 제품조회", type=["xlsx"], key="product_file")
    if list_file and product_file:
        signature = (list_file.name, list_file.size, product_file.name, product_file.size)
        if signature != st.session_state.file_signature:
            try:
                with st.spinner("업로드 파일을 분석하는 중입니다..."):
                    st.session_state.records = process_files(BytesIO(list_file.getvalue()), BytesIO(product_file.getvalue()))
                    st.session_state.file_signature = signature
                st.success(f"분석 완료: 전표 {len(st.session_state.records):,}건")
            except Exception as exc:
                st.session_state.records = None
                st.error(f"파일을 처리하지 못했습니다. 엑셀 형식을 확인해 주세요.\n\n{exc}")
    else:
        st.info("두 개의 .xlsx 파일을 모두 올려 주세요.")

records = st.session_state.records
if records is not None:
    if not records:
        st.warning("소분 또는 대연점 창고에 해당하는 분석 대상 행을 찾지 못했습니다.")
        st.stop()
    frame = display_dataframe(records)
    available_dates = pd.to_datetime(frame["거래일자"], errors="coerce").dropna()
    min_date, max_date = available_dates.min().date(), available_dates.max().date()
    with st.container(border=True):
        st.subheader("조회 조건", divider="gray")
        with st.form("filter_form", border=False):
            start_column, end_column = st.columns(2)
            with start_column:
                start_date = st.date_input("시작일", value=min_date, key="filter_start_date")
            with end_column:
                end_date = st.date_input("종료일", value=max_date, key="filter_end_date")
            keyword = st.text_input("투입품목명", placeholder="예: 삼겹살")
            st.form_submit_button("조회", icon=":material/search:", type="primary")

    if start_date > end_date:
        st.error("시작일은 종료일보다 늦을 수 없습니다.")
        st.stop()
    filtered = frame[(frame["거래일자"] >= start_date.isoformat()) & (frame["거래일자"] <= end_date.isoformat())]
    if keyword:
        filtered = filtered[filtered["투입품목명"].str.contains(keyword, case=False, na=False)]

    metrics = st.columns(4)
    metrics[0].metric("조회 전표", f"{len(filtered):,}건")
    metrics[1].metric("총 투입중량", f"{filtered['투입중량(kg)'].sum():,.2f} kg")
    metrics[2].metric("총 감모중량", f"{filtered['감모수량(kg)'].sum():,.2f} kg")
    metrics[3].metric("총 감모금액", f"{filtered['감모금액(원)'].sum():,.0f}원")

    st.subheader("전표별 분석 결과")
    st.dataframe(
        filtered,
        hide_index=True,
        column_config={
            name: st.column_config.NumberColumn(format="%,.2f") for name in ["투입중량(kg)", "제조중량(kg)", "감모수량(kg)"]
        } | {
            name: st.column_config.NumberColumn(format="%,.0f") for name in ["투입수량", "투입원가(합계)", "제조수량", "납품원가(합계)", "감모금액(원)", "소분작업비용(합계)"]
        } | {
            name: st.column_config.NumberColumn(format="%.2f%%") for name in ["감모비율(%)", "작업비비율(%)"]
        },
        key="results_table",
    )
    st.download_button("결과 엑셀 다운로드", data=to_excel(filtered), file_name="엔포유_소분작업분석.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", icon=":material/download:")

    lookup = {(item["date"], item["doc_no"]): item for item in records}
    choices = [(row["거래일자"], row["전표번호"]) for _, row in filtered.iterrows()]
    if choices:
        st.subheader("전표 상세")
        selected = st.selectbox("상세 전표", choices, format_func=lambda value: f"{value[0]} | 전표번호 {value[1]}")
        detail = lookup[selected]
        cost_metrics = st.columns(3)
        cost_metrics[0].metric("투입원가", f"{detail['inp_cost']:,.0f}원")
        cost_metrics[1].metric("납품원가", f"{detail['out_cost']:,.0f}원")
        cost_metrics[2].metric("소분작업비용", f"{detail['work_cost']:,.0f}원")
        left, right = st.columns(2)
        for column, title, items in ((left, "투입 품목 (소분)", detail["inp_details"]), (right, "산출 품목 (대연점)", detail["out_details"])):
            with column.container(border=True):
                st.markdown(f"#### {title}")
                st.dataframe(pd.DataFrame(items).rename(columns={"code": "코드", "name": "품목명", "qty": "수량", "price": "금액", "unit_weight": "단중량(g)", "weight_kg": "총중량(kg)"}), hide_index=True)
