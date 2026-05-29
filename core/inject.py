"""master.xlsx에 기준값과 업로드 파일 데이터를 주입."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

import openpyxl
from openpyxl.utils import get_column_letter

# 기준 시트 이름 (원본 파일에 끝에 공백 있음)
BASE_SHEET = "배분계산서(기준) "

# 업로드 슬롯 → 대상 시트 이름
UPLOAD_TARGET_SHEETS: Dict[str, str] = {
    "janggeum_salary": "장금급여입력",
    "heungah_salary": "흥아급여입력",
    "janggeum_cost": "장금비용입력",
    "heungah_cost": "흥아비용입력",
    "janggeum_retire": "장금퇴충",
    "heungah_retire": "흥아퇴충",
}


def inject_base_values(
    wb: openpyxl.Workbook,
    year: int,
    quarter: int,
    fx_rate: float,
    janggeum_vat_sales: float,
    heungah_vat_sales: float,
) -> None:
    """배분계산서(기준) 시트의 H2/I2/J2/L2/M2에 값 주입."""
    ws = wb[BASE_SHEET]
    ws["H2"] = int(year)
    ws["I2"] = int(quarter)
    ws["J2"] = float(fx_rate)
    ws["L2"] = float(janggeum_vat_sales)
    ws["M2"] = float(heungah_vat_sales)


def _clear_sheet_data(ws) -> None:
    """시트의 모든 셀 값을 지움 (병합/서식은 유지)."""
    # 병합 해제 → 값 클리어 → 재병합. 여기선 값만 지우므로 병합은 그대로 둠.
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None


def inject_uploaded_excel(
    wb: openpyxl.Workbook,
    target_sheet_name: str,
    uploaded_file_path: Path,
) -> None:
    """업로드한 엑셀의 첫 시트를 값만 추출해서 대상 시트에 시트 전체 붙여넣기.

    - 대상 시트의 기존 값은 모두 지움
    - 업로드 파일의 첫 시트 값만(data_only=True) 가져와서 동일 좌표에 기록
    - 수식 결과가 캐시되어 있지 않으면 값이 None이 될 수 있으므로
      data_only=False 로 한 번 더 시도하고, 수식 셀이면 그 결과값 대신 텍스트를 사용
    """
    if target_sheet_name not in wb.sheetnames:
        raise ValueError(f"대상 시트가 존재하지 않습니다: {target_sheet_name}")

    target_ws = wb[target_sheet_name]

    # 업로드 파일을 '값만' 모드로 로드 (수식은 캐시된 결과값으로 대체됨)
    src_wb_vals = openpyxl.load_workbook(uploaded_file_path, data_only=True)
    src_ws_vals = src_wb_vals[src_wb_vals.sheetnames[0]]

    # 기존 값 클리어
    _clear_sheet_data(target_ws)

    # 값 복사
    max_row = src_ws_vals.max_row or 0
    max_col = src_ws_vals.max_column or 0
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = src_ws_vals.cell(row=r, column=c).value
            if v is not None:
                target_ws.cell(row=r, column=c, value=v)


def populate_workbook(
    template_path: Path,
    output_path: Path,
    *,
    year: int,
    quarter: int,
    fx_rate: float,
    janggeum_vat_sales: float,
    heungah_vat_sales: float,
    uploads: Dict[str, Optional[Path]],
) -> Path:
    """템플릿을 복사 후 모든 입력값/파일을 주입하여 새 워크북으로 저장."""
    wb = openpyxl.load_workbook(template_path, data_only=False, keep_vba=False)

    # 1) 기준 입력값 주입
    inject_base_values(
        wb,
        year=year,
        quarter=quarter,
        fx_rate=fx_rate,
        janggeum_vat_sales=janggeum_vat_sales,
        heungah_vat_sales=heungah_vat_sales,
    )

    # 2) 업로드 파일들을 각 대상 시트에 주입
    for slot, target_sheet in UPLOAD_TARGET_SHEETS.items():
        path = uploads.get(slot)
        if path is None:
            continue
        inject_uploaded_excel(wb, target_sheet, Path(path))

    wb.save(output_path)
    return output_path
