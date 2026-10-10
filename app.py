import io
from collections import defaultdict
from itertools import zip_longest

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROWS = 5  # rows in room-wise sheet


# ---------------------------------------------------------------- logic
def compress(regs):
    """['511524104064','511524104065','511524104067'] -> '511524104064 to 511524104065, 511524104067'"""
    nums = sorted(int(r) for r in regs)
    parts, start, prev = [], nums[0], nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        parts.append(f"{start} to {prev}" if start != prev else str(start))
        start = prev = n
    parts.append(f"{start} to {prev}" if start != prev else str(start))
    return ", ".join(parts)


def allocate(students, rooms):
    """students: DataFrame(reg_no, department, subcode, ...); rooms: list of (room_no, capacity).
    Each room is filled with 2-3 different departments (largest remaining dept first),
    so same-department students are not all together."""
    students = students.sort_values("reg_no").reset_index(drop=True)
    queues = {d: list(g.index) for d, g in students.groupby("department")}
    placed = []  # (room_no, student_index)
    used_rooms = []
    for room, cap in rooms:
        if not any(queues.values()):
            break
        left, used = cap, set()
        while left > 0:
            avail = [d for d in queues if queues[d]]
            if not avail:
                break
            cand = [d for d in avail if d not in used] or avail
            d = max(cand, key=lambda x: len(queues[x]))
            others = [x for x in avail if x != d and x not in used]
            n = (left + 1) // 2 if (others and not used) else left
            n = min(n, len(queues[d]))
            for idx in queues[d][:n]:
                placed.append((room, idx))
            queues[d] = queues[d][n:]
            used.add(d)
            left -= n
        used_rooms.append(room)
    if any(queues.values()):
        raise ValueError("Rooms podhala! Innum rooms add pannunga (rooms.csv).")
    # seat order inside each room: same department should not sit next to each other
    # (above/below in the same column, or left/right in the same row of the sheet)
    final = []
    for room in used_rooms:
        groups = defaultdict(list)
        for r, i in placed:
            if r == room:
                groups[students.loc[i, "department"]].append(i)
        seats = []  # department of every seat position
        order = []
        total = sum(len(v) for v in groups.values())
        for pos in range(total):
            banned = set()
            if pos % ROWS != 0:
                banned.add(seats[pos - 1])
            if pos >= ROWS:
                banned.add(seats[pos - ROWS])
            avail = [d for d in groups if groups[d]]
            ok = [d for d in avail if d not in banned] or avail
            d = max(ok, key=lambda x: len(groups[x]))
            seats.append(d)
            order.append(groups[d].pop(0))
        final += [(room, i) for i in order]
    alloc = students.loc[[i for _, i in final]].copy()
    alloc["room_no"] = [r for r, _ in final]
    return alloc.reset_index(drop=True), used_rooms


def room_grid(regs, rows=ROWS):
    """Column-wise numbering: 5 rows, 1-5 in first column pair, 6-10 in next, ..."""
    cols = -(-len(regs) // rows)
    data = []
    for r in range(rows):
        line = []
        for c in range(cols):
            i = c * rows + r
            line += [i + 1, regs[i]] if i < len(regs) else ["", ""]
        data.append(line)
    return data, cols


def notice_rows(alloc, used_rooms):
    rows = []
    for room in used_rooms:
        sub = alloc[alloc.room_no == room]
        for dept, g in sub.groupby("department", sort=False):
            rows.append({
                "Dept": dept,
                "SubCode": g.subcode.iloc[0],
                "Regno": compress(g.reg_no),
                "Roomno": room,
                "Total": len(g),
            })
    return rows


# ---------------------------------------------------------------- excel
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def make_excel(alloc, used_rooms, title, date, session):
    rows = notice_rows(alloc, used_rooms)
    wb = Workbook()

    ws = wb.active
    ws.title = "Notice Board"
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", bold=True, size=14)
    ws["A2"] = f"DATE: {date}      SESSION: {session}"
    ws["A3"] = f"TOTAL NO OF ROOMS: {len(used_rooms)}      TOTAL NO OF STUDENTS: {len(alloc)}"
    for c in ("A2", "A3"):
        ws[c].font = Font(name="Arial", bold=True)
    head = ["Dept", "SubCode", "Regno", "Roomno", "Total"]
    for j, h in enumerate(head, 1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = Font(name="Arial", bold=True)
        c.border = BOX
    for i, r in enumerate(rows, 6):
        for j, k in enumerate(head, 1):
            c = ws.cell(row=i, column=j, value=r[k])
            c.font = Font(name="Arial")
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDE", (10, 10, 70, 10, 8)):
        ws.column_dimensions[col].width = w

    wa = wb.create_sheet("Allocation")
    cols = ["reg_no", "name", "department", "subcode", "room_no"]
    cols = [c for c in cols if c in alloc.columns]
    wa.append(cols)
    for _, r in alloc.iterrows():
        wa.append([r[c] for c in cols])
    for c in wa[1]:
        c.font = Font(name="Arial", bold=True)
    for col in "ABCDE":
        wa.column_dimensions[col].width = 18

    for room in used_rooms:
        w = wb.create_sheet(str(room)[:31])
        w["A1"] = f"ROOM: {room}"
        w["A1"].font = Font(name="Arial", bold=True, size=14)
        w["A2"] = f"DATE: {date}      SESSION: {session}"
        w.append([])
        regs = list(alloc[alloc.room_no == room].reg_no)
        data, cols = room_grid(regs)
        w.append(["S.No", "Reg No"] * cols)
        for c in w[4]:
            c.font = Font(name="Arial", bold=True)
            c.border = BOX
        for line in data:
            w.append(line)
            for c in w[w.max_row]:
                c.border = BOX
        for k in range(cols):
            w.column_dimensions[chr(65 + 2 * k)].width = 7
            w.column_dimensions[chr(66 + 2 * k)].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- pdf
def _table_style():
    return TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ])


def make_notice_pdf(alloc, used_rooms, title, date, session):
    ss = getSampleStyleSheet()
    small = ss["BodyText"]
    small.fontSize = 8
    small.leading = 10
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
    el = [
        Paragraph(f"<b>{title}</b>", ss["Title"]),
        Paragraph(f"<b>DATE:</b> {date} &nbsp;&nbsp; <b>SESSION:</b> {session}", ss["Normal"]),
        Paragraph(f"<b>TOTAL ROOMS:</b> {len(used_rooms)} &nbsp;&nbsp; <b>TOTAL STUDENTS:</b> {len(alloc)}", ss["Normal"]),
        Spacer(1, 10),
    ]
    data = [["Dept", "SubCode", "Regno", "Roomno", "Total"]]
    for r in notice_rows(alloc, used_rooms):
        data.append([r["Dept"], r["SubCode"], Paragraph(r["Regno"], small), r["Roomno"], r["Total"]])
    t = Table(data, colWidths=[45, 55, 320, 50, 40], repeatRows=1)
    t.setStyle(_table_style())
    el.append(t)
    doc.build(el)
    return buf.getvalue()


def make_room_pdf(alloc, used_rooms, title, date, session):
    ss = getSampleStyleSheet()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), topMargin=40, bottomMargin=40, leftMargin=20, rightMargin=20)
    el = []
    for k, room in enumerate(used_rooms):
        regs = list(alloc[alloc.room_no == room].reg_no)
        el += [
            Paragraph(f"<b>{title}</b>", ss["Heading3"]),
            Paragraph(f"<b>ROOM: {room}</b> &nbsp;&nbsp; DATE: {date} &nbsp;&nbsp; SESSION: {session} &nbsp;&nbsp; TOTAL: {len(regs)}", ss["Normal"]),
            Spacer(1, 12),
        ]
        grid, cols = room_grid(regs)
        data = [["S.No", "Reg No"] * cols] + grid
        t = Table(data, colWidths=[32, 88] * cols, rowHeights=26)
        t.setStyle(_table_style())
        el.append(t)
        if k < len(used_rooms) - 1:
            el.append(PageBreak())
    doc.build(el)
    return buf.getvalue()


# ---------------------------------------------------------------- UI
def read_any(f):
    if f.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(f, dtype=str)
    return pd.read_csv(f, dtype=str)


def main():
    import streamlit as st

    st.set_page_config(page_title="Exam Seating", layout="wide")
    st.title("Exam Room Allotment")

    c1, c2, c3 = st.columns([3, 1, 1])
    title = c1.text_input("Exam title", "ANNA UNIVERSITY THEORY EXAMINATIONS NOV-DEC-2026")
    date = c2.date_input("Date").strftime("%d-%b-%Y")
    session = c3.selectbox("Session", ["FN", "AN"])

    f1 = st.file_uploader("Students file (reg_no, department)", type=["csv", "xlsx"])
    f2 = st.file_uploader("Rooms file (room_no, capacity)", type=["csv", "xlsx"])
    if not (f1 and f2):
        st.info("Rendu file um upload pannunga.")
        return

    students, rooms_df = read_any(f1), read_any(f2)
    for need, df in (("reg_no", students), ("department", students), ("room_no", rooms_df), ("capacity", rooms_df)):
        if need not in df.columns:
            st.error(f"'{need}' column file la illa. Header check pannunga.")
            return
    students["reg_no"] = students["reg_no"].str.strip()

    st.subheader("Subject code (ovvoru department ku)")
    cols = st.columns(4)
    sub = {}
    for i, d in enumerate(sorted(students.department.unique())):
        sub[d] = cols[i % 4].text_input(d, key=f"sub_{d}")
    students["subcode"] = students.department.map(sub)

    if st.button("Allocate", type="primary"):
        rooms = [(r.room_no, int(r.capacity)) for r in rooms_df.itertuples()]
        try:
            alloc, used = allocate(students, rooms)
        except ValueError as e:
            st.error(str(e))
            return
        st.session_state["res"] = (alloc, used)

    if "res" in st.session_state:
        alloc, used = st.session_state["res"]
        st.success(f"{len(alloc)} students, {len(used)} rooms")
        st.dataframe(pd.DataFrame(notice_rows(alloc, used)), use_container_width=True)
        d1, d2, d3 = st.columns(3)
        d1.download_button("Excel (all sheets)", make_excel(alloc, used, title, date, session), "allotment.xlsx")
        d2.download_button("Notice board PDF", make_notice_pdf(alloc, used, title, date, session), "notice_board.pdf")
        d3.download_button("Room wise PDF", make_room_pdf(alloc, used, title, date, session), "room_wise.pdf")


if __name__ == "__main__":
    main()