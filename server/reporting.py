from __future__ import annotations

"""
Reporting helpers extracted from app.py for easier troubleshooting and reuse.
This module renders a monthly expense report PDF from a report dict returned by
the app layer's _build_monthly_report(period, ...).
"""

from datetime import datetime


def render_report_pdf(report: dict) -> bytes:
    """Render a polished Monthly Expense Report PDF with charts and styled sections.
    Expects the report dict with keys: period, currencyBuckets, rows.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        )
        from reportlab.graphics.shapes import Drawing, String, Circle
        from reportlab.graphics.charts.barcharts import HorizontalBarChart
        from reportlab.graphics.charts.piecharts import Pie
    except Exception as e:
        raise RuntimeError(
            f"PDF generation requires reportlab. Install it: pip install reportlab ({e})"
        )

    # Local helpers (no external deps) ---------------------------------------
    def _round2(x):
        try:
            return round(float(x or 0), 2)
        except Exception:
            return 0.0

    def sum_bucket_totals(currency_buckets: dict):
        gross = net = vat = 0.0
        for b in currency_buckets.values():
            t = b.get('totals', {})
            gross += float(t.get('gross', 0) or 0)
            net += float(t.get('net', 0) or 0)
            vat += float(t.get('vat', 0) or 0)
        return _round2(gross), _round2(net), _round2(vat)

    def make_cards_table(doc, styles, cards: list[tuple[str, str]]):
        data = [[Paragraph(f"<b>{title}</b><br/>{value}", styles['BodyText']) for title, value in cards]]
        t = Table(data, colWidths=[(doc.width)/len(cards)]*len(cards))
        t.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#d0d0d0')),
            ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#e0e0e0')),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f7f7f7')),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        return t

    def dict_table(title_left: str, data_dict: dict, col1w=90*mm, col2w=30*mm):
        rows = [[title_left, 'Amount']]
        for k, v in data_dict.items():
            rows.append([str(k), f"{_round2(v):.2f}"])
        t = Table(rows, colWidths=[col1w, col2w])
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eaeaea')),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ]))
        return t

    def bar_chart_from_dict(title: str, data_dict: dict):
        try:
            if not data_dict:
                return None
            labels = list(data_dict.keys())
            values = [float(data_dict[k] or 0) for k in labels]
            h = 140
            w = 340
            d = Drawing(w, h)
            chart = HorizontalBarChart()
            chart.x = 60
            chart.y = 10
            chart.height = h - 30
            chart.width = w - 80
            chart.data = [values]
            chart.categoryAxis.categoryNames = labels
            chart.valueAxis.valueMin = 0
            chart.bars[0].fillColor = colors.HexColor('#5b8cff')
            d.add(chart)
            d.add(String(0, h-10, title, fontSize=10))
            return d
        except Exception:
            return None

    def donut_from_dict(title: str, data_dict: dict):
        try:
            if not data_dict:
                return None
            labels = list(data_dict.keys())
            values = [float(data_dict[k] or 0) for k in labels]
            size = 160
            d = Drawing(size, size)
            p = Pie()
            p.x = 10
            p.y = 10
            p.width = size-20
            p.height = size-20
            p.data = values
            p.labels = [str(l) for l in labels]
            palette = [
                colors.HexColor('#5b8cff'), colors.HexColor('#7b61ff'), colors.HexColor('#22c55e'),
                colors.HexColor('#f59e0b'), colors.HexColor('#ef4444')
            ]
            for i, s in enumerate(p.slices):
                s.fillColor = palette[i % len(palette)]
            d.add(p)
            d.add(Circle(size/2, size/2, 26, fillColor=colors.white, strokeColor=colors.white))
            d.add(String(0, size+2, title, fontSize=10))
            return d
        except Exception:
            return None

    # Build document ----------------------------------------------------------
    from io import BytesIO
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=16*mm, rightMargin=16*mm, topMargin=16*mm, bottomMargin=16*mm
    )
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontSize=18, leading=22)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=14, leading=18)
    h3 = ParagraphStyle('h3', parent=styles['Heading3'], fontSize=12, leading=16)
    normal = styles['BodyText']

    elems: list = []

    # Header
    ts = datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    title = f"Monthly Expense Report — {report.get('period','')}"
    header_table = Table([
        [Paragraph(title, h1), Paragraph('<i>Company Logo</i>', ParagraphStyle('logo', parent=normal, alignment=2))],
        [Paragraph(f"Generated {ts}", normal), '']
    ], colWidths=[doc.width*0.75, doc.width*0.25])
    header_table.setStyle(TableStyle([
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elems.append(header_table)
    elems.append(Spacer(1, 6))

    # Executive Summary
    currency_buckets = report.get('currencyBuckets', {}) or {}
    gross, net, vat = sum_bucket_totals(currency_buckets)
    elems.append(Paragraph('Executive Summary', h2))
    elems.append(make_cards_table(doc, styles, [
        ('Total Gross', f"{gross:.2f}"),
        ('Total Net', f"{net:.2f}"),
        ('Total VAT', f"{vat:.2f}"),
    ]))
    elems.append(Spacer(1, 6))

    # Totals per currency
    if len(currency_buckets) > 0:
        rows = [['Currency', 'Gross', 'Net', 'VAT']]
        for cur, b in currency_buckets.items():
            t = b.get('totals', {})
            rows.append([cur, f"{_round2(t.get('gross')):.2f}", f"{_round2(t.get('net')):.2f}", f"{_round2(t.get('vat')):.2f}"])
        t = Table(rows, colWidths=[30*mm, 30*mm, 30*mm, 30*mm])
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eaeaea')),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ]))
        elems.append(t)
        elems.append(Spacer(1, 8))

    # Prepare rows by currency for per-currency computations
    all_rows = report.get('rows', [])
    rows_by_cur: dict[str, list] = {}
    for r in all_rows:
        rows_by_cur.setdefault(r.get('currency') or 'Unknown', []).append(r)

    # Breakdowns per currency
    elems.append(Paragraph('Breakdowns', h2))
    for cur, b in currency_buckets.items():
        elems.append(Paragraph(f'Currency: {cur}', h3))
        by_cat = b.get('byCategory', {})
        by_pay = b.get('byPaymentMethod', {})

        # Charts
        bar = bar_chart_from_dict('By Category', by_cat)
        donut = donut_from_dict('By Payment Method', by_pay)
        charts_row = []
        if bar:
            charts_row.append(bar)
        if donut:
            charts_row.append(donut)
        if charts_row:
            cw = [doc.width/len(charts_row)]*len(charts_row)
            charts_table = Table([charts_row], colWidths=cw)
            charts_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ]))
            elems.append(charts_table)
            elems.append(Spacer(1, 4))

        # Tables
        elems.append(dict_table('Category', by_cat))
        elems.append(Spacer(1, 4))
        elems.append(dict_table('Payment Method', by_pay))
        elems.append(Spacer(1, 6))

        # Company card overview (Done only)
        cc_total = 0.0
        for r in rows_by_cur.get(cur, []):
            if (r.get('status') or '').lower() == 'done' and (r.get('paymentMethod') or '').lower().startswith('company'):
                cc_total += float(r.get('gross') or 0)
        elems.append(make_cards_table(doc, styles, [('Company Card Spent', f"{_round2(cc_total):.2f}")]))

        # Reimbursements by employee (Done + Personal/Cash)
        reimb: dict[str, float] = {}
        for r in rows_by_cur.get(cur, []):
            pm = (r.get('paymentMethod') or '').lower()
            st = (r.get('status') or '').lower()
            if st == 'done' and (pm == 'personal' or pm == 'cash'):
                key = (r.get('payer') or 'Unknown').strip() or 'Unknown'
                reimb[key] = reimb.get(key, 0.0) + float(r.get('gross') or 0)
        if reimb:
            elems.append(Spacer(1, 4))
            elems.append(Paragraph('Reimbursements owed (Done, Personal/Cash)', h3))
            elems.append(dict_table('Employee', reimb))
        elems.append(Spacer(1, 6))

    # Pending overview (aggregated)
    ip_count = ur_count = 0
    ip_gross = ur_gross = 0.0
    for b in currency_buckets.values():
        p = b.get('pending', {})
        ip = p.get('inProgress', { 'count': 0, 'gross': 0 })
        ur = p.get('underReview', { 'count': 0, 'gross': 0 })
        ip_count += int(ip.get('count', 0) or 0)
        ur_count += int(ur.get('count', 0) or 0)
        ip_gross += float(ip.get('gross', 0) or 0)
        ur_gross += float(ur.get('gross', 0) or 0)
    elems.append(Paragraph('Pending Overview', h2))
    elems.append(make_cards_table(doc, styles, [
        ('In‑Progress', f"{ip_count} • {ip_gross:.2f}"),
        ('Under Review', f"{ur_count} • {ur_gross:.2f}"),
    ]))
    elems.append(Spacer(1, 8))

    # Detailed rows
    rows = report.get('rows', [])
    elems.append(Paragraph('Detailed Rows', h2))
    if not rows:
        elems.append(Paragraph('No rows for this period with the selected statuses.', normal))
    else:
        header = ['Date', 'Payer', 'Category', 'Payment', 'Gross', 'Net', 'VAT', 'Currency', 'Status']
        data = [header]
        for r in rows:
            data.append([
                r.get('date',''), r.get('payer',''), r.get('category',''), r.get('paymentMethod',''),
                f"{_round2(r.get('gross')):.2f}", f"{_round2(r.get('net')):.2f}", f"{_round2(r.get('vat')):.2f}",
                r.get('currency',''), r.get('status',''),
            ])
        table = Table(data, repeatRows=1)
        style_cmds = [
            ('GRID', (0,0), (-1,-1), 0.4, colors.lightgrey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eaeaea')),
            ('ALIGN', (4,1), (6,-1), 'RIGHT'),
        ]
        for i in range(1, len(data)):
            bg = colors.whitesmoke if i % 2 == 0 else colors.white
            style_cmds.append(('BACKGROUND', (0,i), (-1,i), bg))
            status_text = str(data[i][-1] or '').lower()
            status_col = colors.HexColor('#22c55e') if 'done' in status_text else (
                colors.HexColor('#f59e0b') if 'progress' in status_text else colors.HexColor('#ef4444')
            )
            style_cmds.append(('BACKGROUND', (-1,i), (-1,i), status_col))
            style_cmds.append(('TEXTCOLOR', (-1,i), (-1,i), colors.white))
        table.setStyle(TableStyle(style_cmds))
        elems.append(table)

    # Footer
    def _footer(canvas, _doc):
        canvas.saveState()
        footer_text = 'In‑House Expensify — Confidential • Page %d' % (_doc.page)
        canvas.setFillColor(colors.HexColor('#666666'))
        canvas.setFont('Helvetica', 9)
        canvas.drawString(16*mm, 12*mm, footer_text)
        canvas.restoreState()

    doc.build(elems, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


def render_error_pdf(message: str) -> bytes:
    """Return a small error PDF describing the failure. Always returns a valid PDF if reportlab is present.
    If reportlab is missing, returns plain text bytes as a last resort.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.units import mm
    except Exception:
        return (message or 'Error').encode('utf-8')
    from io import BytesIO
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=16*mm, rightMargin=16*mm, topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    elems = [Paragraph('Report Generation Error', styles['Heading2']), Paragraph(message or 'Unknown error', styles['BodyText'])]
    doc.build(elems)
    return buf.getvalue()

