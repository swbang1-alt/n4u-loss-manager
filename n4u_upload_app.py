```python
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


# =========================================================
# 화면 표시 컬럼
# =========================================================
DISPLAY_COLUMNS = [
    "거래일자",
    "전표번호",
    "투입품목코드",
    "투입품목명",
    "투입수량",
    "투입중량(kg)",
    "투입원가",
    "투입원가(합계)",
    "산출품목코드",
    "산출품목명",
    "제조수량",
    "제조중량(kg)",
    "납품원가",
    "납품원가(합계)",
    "감모수량(kg)",
    "감모금액(합계)",
    "감모비율(%)",
    "작업비비율(%)",
    "소분작업비용(합계)",
]


# =========================================================
# 공통 함수
# =========================================================
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

    cleaned = re.sub(
        r"[^0-9.+-]",
        "",
        str(value).replace(",", ""),
    )

    try:
        return float(cleaned)
    except ValueError:
        return default


# =========================================================
# 엑셀 컬럼 찾기
# =========================================================
def find_columns(
    rows: list[tuple[Any, ...]],
) -> tuple[int, dict[str, int]]:
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

            header = str(
                cell or ""
            ).replace(" ", "").lower()

            for key, words in patterns.items():

                if (
                    key not in found
                    and any(
                        word in header
                        for word in words
                    )
                ):
                    found[key] = col_index

            # 원본 n4u.py와 동일하게 작업비용에는 판매금액을 쓴다.
            # 판매금액이 없는 옛 양식만 합계금액, 공급가액 순으로 대체한다.
            if header == "판매금액":

                found["price"] = col_index

            elif (
                header == "합계금액"
                and "price" not in found
            ):

                found["price"] = col_index

            elif (
                header == "공급가액"
                and "price" not in found
            ):

                found["price"] = col_index

        if (
            "warehouse" in found
            and "price" in found
        ):
            return row_index, found

    raise ValueError(
        "소분리스트에서 필수 열인 "
        "'창고명'과 '판매금액/금액'을 찾지 못했습니다."
    )


# =========================================================
# 날짜 추출
# =========================================================
def extract_date(value: Any) -> str:

    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    match = re.search(
        r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}",
        str(value or ""),
    )

    return (
        match.group()
        .replace(".", "-")
        .replace("/", "-")
        if match
        else ""
    )


# =========================================================
# 값 검색
# =========================================================
def first_match(
    row: tuple[Any, ...],
    pattern: str,
) -> str:

    for cell in row:

        match = re.search(
            pattern,
            str(cell or ""),
        )

        if match:
            return match.group()

    return ""


# =========================================================
# 제품 중량 읽기
# =========================================================
def read_product_weights(
    file: BinaryIO,
) -> dict[str, float]:

    workbook = openpyxl.load_workbook(
        file,
        data_only=True,
        read_only=True,
    )

    sheet = workbook.active

    weights: dict[str, float] = {}

    for row in sheet.iter_rows(
        values_only=True
    ):

        if len(row) <= 26:
            continue

        code = normalize_code(
            row[1]
        )  # B: 제품코드

        weight = to_number(
            row[26],
            default=-1,
        )  # AA: 중량

        if code and weight >= 0:
            weights[code] = weight

    return weights


# =========================================================
# 파일 분석
# =========================================================
def process_files(
    list_file: BinaryIO,
    product_file: BinaryIO,
) -> list[dict[str, Any]]:

    weights = read_product_weights(
        product_file
    )

    workbook = openpyxl.load_workbook(
        list_file,
        data_only=True,
        read_only=True,
    )

    rows = list(
        workbook.active.iter_rows(
            values_only=True
        )
    )

    header_index, columns = find_columns(
        rows
    )

    grouped: dict[
        tuple[str, str],
        dict[str, Any],
    ] = {}

    current_doc = "-"
    cu
```
