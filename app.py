# -*- coding: utf-8 -*-
"""
CIC LOGIS BOOKING CONSOLE (모바일 웹/PC 겸용 부킹 리스트 관리 시스템)
"""

import os
import re
import csv
import io
import datetime
import urllib.request
import urllib.error
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# PDF 생성 라이브러리 연동
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

st.set_page_config(page_title="CIC LOGIS 부킹 콘솔", page_icon="📦", layout="wide")

THEME_COLOR = "#0F172A"

CLIENT_COLOR_MAP = {
    "TLC": "#D0E1ED",
    "알프스": "#D4EDDA",
    "서하": "#F8D7DA"
}

FORWARDER_COLOR_MAP = {
    "우정": "#D1ECF1",
    "코스모": "#FFE8D6",
    "국양": "#D6C1B3",
    "트랜스올": "#D4EDDA",
    "바롬": "#E2D9F3"
}

COMPANY_INFO = {
    "NAME_EN": "CIC LOGIS CO.,LTD.",
    "NAME_KR": "(주)씨아이씨로지스틱",
    "ADDR_EN": "D4-block, 95, Jayumuyeok-ro, Jung-gu, Incheon, Korea",
    "ADDR_KR": "인천광역시 중구 자유무역로 95 D4블록 301호",
    "TEL": "+82-32-744-4221~6",
    "FAX": "+82-70-4015-5833",
    "EMAIL": "ciclogis@ciclogis.co.kr"
}

def get_pdf_korean_font():
    font_paths = [
        "C:/Windows/Fonts/malgun.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc"
    ]
    for p in font_paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("KoreanFont", p))
                return "KoreanFont"
            except Exception:
                continue
    return "Helvetica"

def draw_cic_letterhead_background(canvas, doc):
    canvas.saveState()
    w, h = doc.pagesize

    foot_x = 20
    foot_y = 16
    canvas.setFont("Helvetica-Bold", 7.5)
    canvas.setFillColor(colors.HexColor("#2563EB"))
    canvas.drawString(foot_x, foot_y + 18, COMPANY_INFO["NAME_EN"])

    kor_font = get_pdf_korean_font()
    canvas.setFont(kor_font, 6.5)
    canvas.setFillColor(colors.HexColor("#F43F5E"))
    canvas.drawString(foot_x, foot_y + 10, COMPANY_INFO["ADDR_EN"] + " / " + COMPANY_INFO["ADDR_KR"])
    canvas.drawString(foot_x, foot_y + 2, "Tel: " + COMPANY_INFO["TEL"] + " | Fax: " + COMPANY_INFO["FAX"] + " | Email: " + COMPANY_INFO["EMAIL"])

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawRightString(w - 20, foot_y + 2, "Page " + str(canvas.getPageNumber()) + " | CIC LOGIS BOOKING CONSOLE")

    canvas.restoreState()

st.markdown(
    '<div style="background-color:' + THEME_COLOR + '; padding:15px; border-radius:8px; color:white; margin-bottom:20px;">'
    '<h3 style="margin:0;">CIC LOGIS BOOKING CONSOLE (모바일/PC 겸용)</h3>'
    '<p style="margin:5px 0 0 0; font-size:13px; color:#94A3B8;">부킹 리스트 전용 관리 및 파일 저장 시스템</p>'
    '</div>',
    unsafe_allow_html=True
)

if "booking_records" not in st.session_state:
    st.session_state.booking_records = []

with st.sidebar:
    st.header("데이터 불러오기")
    fetch_mode = st.radio("방식 선택", ["구글 시트 연동", "엑셀 파일 업로드"])
    
    sheet_url = ""
    sheet_name = "시트1"
    uploaded_file = None

    if fetch_mode == "구글 시트 연동":
        sheet_url = st.text_input("구글 시트 주소 (URL)")
        sheet_name = st.text_input("시트 탭 이름", value="시트1")
        fetch_btn = st.button("구글시트 가져오기", use_container_width=True)
    else:
        uploaded_file = st.file_uploader("엑셀 파일 선택 (.xlsx)", type=["xlsx", "xls"])
        fetch_btn = st.button("엑셀 파일 열기", use_container_width=True)

    st.divider()
    st.header("내보내기 설정")
    use_highlight = st.checkbox("형광칠 사용", value=True)
    use_magnifier = st.checkbox("돋보기 (글자 확대)", value=False)

if fetch_btn:
    rows_data = []
    if fetch_mode == "구글 시트 연동" and sheet_url:
        match_id = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_url)
        if match_id:
            sheet_id = match_id.group(1)
            encoded_sheet_name = urllib.parse.quote(sheet_name)
            csv_url = "https://docs.google.com/spreadsheets/d/" + sheet_id + "/gviz/tq?tqx=out:csv&sheet=" + encoded_sheet_name
            try:
                req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=12) as response:
                    raw_bytes = response.read()
                try:
                    csv_text = raw_bytes.decode('utf-8-sig')
                except UnicodeDecodeError:
                    csv_text = raw_bytes.decode('cp949', errors='ignore')
                reader = csv.reader(io.StringIO(csv_text))
                rows_data = list(reader)
                st.success("구글 시트 연동 성공!")
            except Exception as e:
                st.error("구글 시트 로드 실패: " + str(e))
        else:
            st.error("올바른 구글 시트 주소가 아닙니다.")
    elif fetch_mode == "엑셀 파일 업로드" and uploaded_file:
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            ws = wb.active
            for r in range(1, ws.max_row + 1):
                rows_data.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
            st.success("엑셀 파일 로드 성공!")
        except Exception as e:
            st.error("엑셀 파일 로드 실패: " + str(e))

    if rows_data:
        records = []
        header_idx = -1
        for idx, row in enumerate(rows_data[:12]):
            row_str = " ".join([str(c) for c in row if c is not None]).upper()
            if "MAWB" in row_str or "DEST" in row_str or "업체" in row_str:
                header_idx = idx
                break

        start_row = header_idx + 1 if header_idx != -1 else 0
        
        for row in rows_data[start_row:]:
            if not any(row):
                continue
            vals = [str(c).strip() if c is not None else "" for c in row]
            while len(vals) < 22:
                vals.append("")

            client_val = vals[0] if len(vals) > 0 else ""
            no_val = vals[1] if vals[1].isdigit() else str(len(records) + 1)
            mawb_val = vals[2] if len(vals) > 2 else ""
            dest_val = vals[3] if len(vals) > 3 else ""
            flt_val = vals[4] if len(vals) > 4 else ""
            onboard_val = vals[5] if len(vals) > 5 else ""
            time_val = vals[6] if len(vals) > 6 else ""
            net_val = vals[13] if len(vals) > 13 else ""
            freight_val = vals[15] if len(vals) > 15 else ""
            forwarder_val = vals[12] if len(vals) > 12 else ""

            if not mawb_val and not flt_val and not freight_val:
                continue
            if "MAWB" in mawb_val.upper() or "청구운임" in freight_val or "등록" in freight_val or "NET운임" in net_val.upper():
                continue

            bg_color = CLIENT_COLOR_MAP.get(client_val, None)
            fwd_color = FORWARDER_COLOR_MAP.get(forwarder_val, None)

            records.append({
                "client": client_val,
                "no": no_val,
                "mawb": mawb_val,
                "dest": dest_val,
                "flt": flt_val,
                "onboard": onboard_val,
                "time": time_val,
                "net": net_val,
                "freight": freight_val,
                "forwarder": forwarder_val,
                "bg_color": bg_color,
                "forwarder_bg_color": fwd_color
            })
        st.session_state.booking_records = records

records = st.session_state.booking_records
if records:
    st.subheader("부킹 리스트 관리 (총 " + str(len(records)) + "건)")
    
    all_clients = sorted(list(set(r["client"] for r in records)))
    selected_clients = st.multiselect("업체 필터 (비워두면 전체)", options=all_clients, default=all_clients)
    
    filtered_records = [r for r in records if not selected_clients or r["client"] in selected_clients]

    table_data = []
    for r in filtered_records:
        table_data.append({
            "업체": r["client"],
            "NO": r["no"],
            "MAWB NO": r["mawb"],
            "DEST": r["dest"],
            "FLT NO": r["flt"],
            "Onboard": r["onboard"],
            "Time": r["time"],
            "NET운임": r["net"],
            "청구운임": r["freight"],
            "콘솔사": r["forwarder"]
        })
    
    st.dataframe(table_data, use_container_width=True)

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        if st.button("엑셀 파일로 저장", use_container_width=True):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "부킹리스트"
            thin_border = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                                 top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

            headers = ["업체", "NO", "MAWB NO", "DEST", "FLT NO", "Onboard", "Time", "NET운임", "청구운임", "콘솔사"]
            widths = [14, 8, 18, 10, 12, 12, 45, 12, 14, 14]

            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
            ws['A1'] = "CIC LOGIS | 수출팀 부킹 리스트 (BOOKING LIST)"
            ws['A1'].font = Font(name="맑은 고딕", size=14, bold=True, color="FFFFFF")
            ws['A1'].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 32

            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
                cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
            ws.row_dimensions[2].height = 26

            for r_idx, r in enumerate(filtered_records, start=3):
                row_vals = [r["client"], r["no"], r["mawb"], r["dest"], r["flt"], r["onboard"], r["time"], r["net"], r["freight"], r["forwarder"]]
                
                client_fill = None
                if use_highlight and r.get("bg_color"):
                    hex_c = r["bg_color"].replace("#", "")
                    client_fill = PatternFill(start_color=hex_c, end_color=hex_c, fill_type="solid")

                fwd_fill = None
                if use_highlight and r.get("forwarder_bg_color"):
                    hex_f = r["forwarder_bg_color"].replace("#", "")
                    fwd_fill = PatternFill(start_color=hex_f, end_color=hex_f, fill_type="solid")

                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = Font(name="맑은 고딕", size=10)
                    cell.border = thin_border
                    
                    if c_idx == 10 and fwd_fill:
                        cell.fill = fwd_fill
                    elif client_fill:
                        cell.fill = client_fill
                    
                    is_time_col = (c_idx == 7)
                    cell.alignment = Alignment(horizontal="center" if c_idx in [1, 2, 4, 5, 6] else ("left" if is_time_col or c_idx == 3 else "right"), vertical="center", wrap_text=is_time_col)
                ws.row_dimensions[r_idx].height = 45

            for c_idx, w_val in enumerate(widths, start=1):
                ws.column_dimensions[get_column_letter(c_idx)].width = w_val

            excel_filename = "CIC_LOGIS_부킹리스트_" + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + ".xlsx"
            wb.save(excel_filename)
            
            with open(excel_filename, "rb") as f:
                st.download_button(label="엑셀 파일 다운로드 클릭", data=f, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with col2:
        if HAS_REPORTLAB and st.button("PDF 파일로 저장", use_container_width=True):
            pdf_filename = "CIC_LOGIS_부킹리스트_" + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + ".pdf"
            kor_font = get_pdf_korean_font()
            doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=40, bottomMargin=40)
            elements = []

            title_fs = 18 if use_magnifier else 15
            header_fs = 11 if use_magnifier else 8.5
            cell_fs = 10 if use_magnifier else 8
            cell_leading = 13 if use_magnifier else 10
            row_padding = 10 if use_magnifier else 6

            title_style = ParagraphStyle('TitleStyle', fontName=kor_font, fontSize=title_fs, textColor=colors.HexColor("#0F172A"), alignment=1, spaceAfter=14)
            elements.append(Paragraph("<b>CIC LOGIS 수출팀 부킹 리스트 (BOOKING LIST)</b>", title_style))

            all_meta = [
                ("client", "업체", 70), ("no", "NO", 35), ("mawb", "MAWB NO", 110),
                ("dest", "DEST", 55), ("flt", "FLT NO", 70), ("onboard", "Onboard", 70),
                ("time", "Time", 185), ("net", "NET운임", 75), ("freight", "청구운임", 75),
                ("forwarder", "콘솔사", 75)
            ]

            cell_style = ParagraphStyle('Cell', fontName=kor_font, fontSize=cell_fs, leading=cell_leading, textColor=colors.HexColor("#0F172A"))
            cell_center = ParagraphStyle('CellCenter', parent=cell_style, alignment=1)
            cell_right = ParagraphStyle('CellRight', parent=cell_style, alignment=2)
            header_style = ParagraphStyle('Header', fontName=kor_font, fontSize=header_fs, leading=header_fs+3, textColor=colors.white, alignment=1)

            headers = [Paragraph(m[1], header_style) for m in all_meta]
            table_data = [headers]
            table_style_cmds = [
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ('BOTTOMPADDING', (0, 0), (-1, -1), row_padding),
                ('TOPPADDING', (0, 0), (-1, -1), row_padding),
            ]

            for r_idx, r in enumerate(filtered_records, start=1):
                row_p = [
                    Paragraph(r["client"], cell_center),
                    Paragraph(r["no"], cell_center),
                    Paragraph(r["mawb"], cell_center),
                    Paragraph(r["dest"], cell_center),
                    Paragraph(r["flt"], cell_center),
                    Paragraph(r["onboard"], cell_center),
                    Paragraph(r["time"], cell_style),
                    Paragraph(r["net"], cell_right),
                    Paragraph(r["freight"], cell_right),
                    Paragraph(r["forwarder"], cell_center)
                ]
                table_data.append(row_p)

                if use_highlight:
                    bg_c = r.get("bg_color")
                    if bg_c:
                        try:
                            table_style_cmds.append(('BACKGROUND', (0, r_idx), (8, r_idx), colors.HexColor(bg_c)))
                        except Exception:
                            pass
                    fwd_c = r.get("forwarder_bg_color")
                    if fwd_c:
                        try:
                            table_style_cmds.append(('BACKGROUND', (9, r_idx), (9, r_idx), colors.HexColor(fwd_c)))
                        except Exception:
                            pass

            col_w = [m[2] for m in all_meta]
            pdf_table = Table(table_data, colWidths=col_w, repeatRows=1)
            pdf_table.setStyle(TableStyle(table_style_cmds))
            elements.append(pdf_table)

            doc.build(elements, onFirstPage=draw_cic_letterhead_background, onLaterPages=draw_cic_letterhead_background)

            with open(pdf_filename, "rb") as f:
                st.download_button(label="PDF 파일 다운로드 클릭", data=f, file_name=pdf_filename, mime="application/pdf", use_container_width=True)
else:
    st.info("왼쪽 사이드바에서 구글 시트 주소를 넣거나 엑셀 파일을 업로드해 주세요.")