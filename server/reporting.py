from __future__ import annotations

'''PDF reporting utilities for the Monthly Expense Report.

This module converts the raw report payload returned by `_build_monthly_report`
into a polished, multi-page PDF optimised for executives and accountants.
All monetary values are converted to CHF before aggregation and rendering.
'''

import json
import os
from datetime import datetime, date
from functools import lru_cache
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple
from urllib.request import urlopen

LOGO_URL = (
    'https://images.squarespace-cdn.com/content/v1/67fcd0d3a9d60d2152d4fe76/'
    '7914380b-9930-4ffd-a059-472dfb9cc664/Bild45.png?format=2500w'
)
FONT_DIR = Path(__file__).resolve().parent / 'assets' / 'fonts'
FONT_REGULAR_PATH = FONT_DIR / 'Inter-Regular.ttf'
FONT_SEMIBOLD_PATH = FONT_DIR / 'Inter-SemiBold.ttf'
SUMMARY_FALLBACK = 'Summary unavailable. See metrics above.'


# ---------------------------------------------------------------------------
# Helper functions exposed for reuse + unit tests
# ---------------------------------------------------------------------------
def formatCHF(value: float | int | str | None) -> str:
    '''Format a value as CHF with two decimals and Swiss thousands separators.'''
    try:
        amount = float(value or 0)
    except Exception:
        amount = 0.0
    formatted = f"{amount:,.2f}".replace(',', "'")
    return f'CHF {formatted}'


def percent(part: float | int | None, total: float | int | None) -> str:
    '''Return the share of ``part`` over ``total`` as a percentage string.'''
    try:
        p = float(part or 0)
        t = float(total or 0)
        if t == 0:
            return '0.0%'
        return f'{(p / t) * 100:.1f}%'
    except Exception:
        return '0.0%'


def _normalize_currency(code: str | None) -> str:
    if not code:
        return 'CHF'
    return str(code).strip().upper() or 'CHF'


def _safe_float(value: Any) -> float:
    try:
        return round(float(value or 0), 2)
    except Exception:
        return 0.0


def _describe_rates(rates: Dict[str, float]) -> str:
    if not rates:
        return 'Standard 1:1 CHF conversion'
    parts = []
    for cur in sorted(rates.keys()):
        try:
            parts.append(f"1 {cur} = {float(rates[cur]):.4f} CHF")
        except Exception:
            continue
    return 'Internal fixed conversion policy: ' + ', '.join(parts)


def buildTables(report: Dict[str, Any]) -> Dict[str, Any]:
    '''Aggregate CHF metrics and table data from the raw report payload.'''
    rows: List[Dict[str, Any]] = list(report.get('rows') or [])
    fx_rates_raw = report.get('fxRatesCHF') or report.get('fxRates') or {}
    fx_rates = {_normalize_currency(k): float(v) for k, v in fx_rates_raw.items() if v is not None}
    if not fx_rates:
        fx_rates = {'CHF': 1.0}

    rows_chf: List[Dict[str, Any]] = []
    total_gross = total_net = total_vat = 0.0
    company_card = 0.0
    reimb_by_employee: Dict[str, float] = {}
    by_category: Dict[str, float] = {}
    by_payment: Dict[str, float] = {}
    pending_ip = {'count': 0, 'amount': 0.0}
    pending_ur = {'count': 0, 'amount': 0.0}

    for row in rows:
        status = (row.get('status') or '').strip() or 'Done'
        status_key = status.lower()
        payment = (row.get('paymentMethod') or 'Other').strip() or 'Other'
        category = (row.get('category') or 'Uncategorised').strip() or 'Uncategorised'
        payer = (row.get('payer') or 'Unknown').strip() or 'Unknown'
        original_currency = _normalize_currency(row.get('currency') or row.get('originalCurrency') or 'CHF')
        fx_rate = float(fx_rates.get(original_currency, fx_rates.get('CHF', 1.0)) or 0) or 1.0

        gross_original = _safe_float(row.get('gross') if row.get('gross') is not None else row.get('originalAmount'))
        net_original = _safe_float(row.get('net'))
        vat_original = _safe_float(row.get('vat'))

        gross_chf = round(gross_original * fx_rate, 2)
        net_chf = round(net_original * fx_rate, 2)
        vat_chf = round(vat_original * fx_rate, 2)

        rows_chf.append({
            'date': row.get('date') or '',
            'payer': payer,
            'category': category,
            'paymentMethod': payment,
            'status': status,
            'amountCHF': gross_chf,
            'netCHF': net_chf,
            'vatCHF': vat_chf,
            'originalAmount': gross_original,
            'originalCurrency': original_currency,
            'receiptUrl': row.get('receiptUrl') or None,
        })

        if status_key == 'in-progress':
            pending_ip['count'] += 1
            pending_ip['amount'] += gross_chf
            continue
        if status_key == 'under review':
            pending_ur['count'] += 1
            pending_ur['amount'] += gross_chf
            continue

        total_gross += gross_chf
        total_net += net_chf
        total_vat += vat_chf

        by_category[category] = by_category.get(category, 0.0) + gross_chf
        by_payment[payment] = by_payment.get(payment, 0.0) + gross_chf

        if payment.lower() == 'company card':
            company_card += gross_chf
        if payment.lower() in {'personal', 'cash'}:
            reimb_by_employee[payer] = reimb_by_employee.get(payer, 0.0) + gross_chf

    rows_chf.sort(key=lambda r: (r.get('date') or '', r.get('payer') or ''))

    def _sorted_table(data: Dict[str, float]) -> List[Tuple[str, float]]:
        return [(k, round(data[k], 2)) for k in sorted(data.keys(), key=lambda x: x.lower())]

    cat_table = _sorted_table(by_category)
    pay_table = _sorted_table(by_payment)
    reimb_table = _sorted_table(reimb_by_employee)
    top_owed = sorted(reimb_table, key=lambda item: item[1], reverse=True)

    top_category = max(cat_table, key=lambda item: item[1]) if cat_table else None

    pending_summary = {
        'inProgress': {'count': pending_ip['count'], 'amount': round(pending_ip['amount'], 2)},
        'underReview': {'count': pending_ur['count'], 'amount': round(pending_ur['amount'], 2)},
    }
    pending_total_amount = pending_summary['inProgress']['amount'] + pending_summary['underReview']['amount']
    pending_total_count = pending_summary['inProgress']['count'] + pending_summary['underReview']['count']

    totals = {
        'grossCHF': round(total_gross, 2),
        'netCHF': round(total_net, 2),
        'vatCHF': round(total_vat, 2),
        'companyCardSpentCHF': round(company_card, 2),
        'reimbursementsOwedCHF': round(sum(amount for _, amount in reimb_table), 2),
        'pendingTotalCHF': round(pending_total_amount, 2),
        'pendingTotalCount': pending_total_count,
    }

    return {
        'rowsCHF': rows_chf,
        'totals': totals,
        'byCategory': cat_table,
        'byPaymentMethod': pay_table,
        'reimbursements': reimb_table,
        'pending': pending_summary,
        'topOwed': top_owed,
        'topCategory': top_category,
        'ratePolicy': report.get('fxPolicy') or _describe_rates(fx_rates),
        'fxRates': fx_rates,
    }


def summarizeForAI(metrics: Dict[str, Any], client: Any | None = None) -> str:
    '''Generate a deterministic 3–5 sentence summary via OpenAI.'''
    payload = {
        'period': metrics.get('period'),
        'totals': metrics.get('totals', {}),
        'topCategory': metrics.get('topCategory'),
        'companyCardSpentCHF': metrics.get('totals', {}).get('companyCardSpentCHF', 0.0),
        'reimbursementsTop': metrics.get('topOwed', [])[:3],
        'pending': {
            'inProgressCount': metrics.get('pending', {}).get('inProgress', {}).get('count', 0),
            'inProgressCHF': metrics.get('pending', {}).get('inProgress', {}).get('amount', 0.0),
            'underReviewCount': metrics.get('pending', {}).get('underReview', {}).get('count', 0),
            'underReviewCHF': metrics.get('pending', {}).get('underReview', {}).get('amount', 0.0),
        },
    }

    try:
        content = json.dumps(payload, ensure_ascii=False)
        if client is None:
            from openai import OpenAI  # type: ignore

            api_key = os.getenv('OPENAI_API_KEY')
            if not api_key:
                raise RuntimeError('Missing OPENAI_API_KEY')
            client = OpenAI(api_key=api_key)

        response = client.chat.completions.create(
            model='gpt-4o-mini',
            temperature=0,
            messages=[
                {
                    'role': 'system',
                    'content': 'You summarize monthly expense reports for executives. '
                               'Output 3-5 short sentences. Use CHF currency. Be neutral and audit-safe.',
                },
                {'role': 'user', 'content': content},
            ],
        )
        text = (response.choices[0].message.content or '').strip()
        return text or SUMMARY_FALLBACK
    except Exception:
        return SUMMARY_FALLBACK


# ---------------------------------------------------------------------------
# Internal helpers for PDF assembly
# ---------------------------------------------------------------------------
ACCENT_COLOR = '#2563EB'
BACKGROUND_COLOR = '#F5F7FB'
BODY_TEXT_COLOR = '#1F2933'
LIGHT_BORDER = '#D9E0EB'
STATUS_COLORS = {
    'done': '#16A34A',
    'in-progress': '#F59E0B',
    'under review': '#EF4444',
}


def _register_fonts() -> Tuple[str, str]:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    base_font = 'Helvetica'
    bold_font = 'Helvetica-Bold'

    try:
        if FONT_REGULAR_PATH.exists():
            pdfmetrics.registerFont(TTFont('Inter', str(FONT_REGULAR_PATH)))
            base_font = 'Inter'
        if FONT_SEMIBOLD_PATH.exists():
            pdfmetrics.registerFont(TTFont('Inter-Bold', str(FONT_SEMIBOLD_PATH)))
            bold_font = 'Inter-Bold'
    except Exception:
        base_font = 'Helvetica'
        bold_font = 'Helvetica-Bold'

    return base_font, bold_font


@lru_cache(maxsize=1)
def _logo_bytes() -> bytes | None:
    try:
        with urlopen(LOGO_URL, timeout=10) as resp:  # nosec: trusted URL from spec
            return resp.read()
    except Exception:
        return None


def _period_bounds(period: str) -> Tuple[date, date]:
    from calendar import monthrange

    try:
        year, month = period.split('-')
        year_i = int(year)
        month_i = int(month)
    except Exception:
        today = datetime.utcnow().date()
        start = date(today.year, today.month, 1)
        end = start
        return start, end

    start = date(year_i, month_i, 1)
    end_day = monthrange(year_i, month_i)[1]
    end = date(year_i, month_i, end_day)
    return start, end


def _load_styles(base_font: str, bold_font: str):
    from reportlab.lib.styles import ParagraphStyle, StyleSheet1

    styles = StyleSheet1()
    styles.add(ParagraphStyle(
        name='HeadingLarge',
        fontName=bold_font,
        fontSize=16,
        leading=20,
        textColor=BODY_TEXT_COLOR,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name='HeadingMedium',
        fontName=bold_font,
        fontSize=13,
        leading=17,
        textColor=BODY_TEXT_COLOR,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name='HeadingSmall',
        fontName=bold_font,
        fontSize=11.5,
        leading=15,
        textColor=BODY_TEXT_COLOR,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name='Body',
        fontName=base_font,
        fontSize=10.5,
        leading=14,
        textColor=BODY_TEXT_COLOR,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name='BodySmall',
        fontName=base_font,
        fontSize=9.2,
        leading=12,
        textColor=BODY_TEXT_COLOR,
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        name='Caption',
        fontName=base_font,
        fontSize=9,
        leading=11,
        textColor='#6B7280',
        spaceAfter=2,
    ))
    return styles


def _metric_cards(cards: List[Tuple[str, str]], doc_width: float, styles, columns: int = 4):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle

    col_width = doc_width / columns
    data: List[List[Any]] = []
    row: List[Any] = []
    for idx, (title, value) in enumerate(cards, start=1):
        para = Paragraph(
            (
                f"<para align=\"left\"><font name=\"{styles['HeadingSmall'].fontName}\" size=\"10.5\">{title}</font><br/>"
                f"<font name=\"{styles['HeadingLarge'].fontName}\" size=\"13\">{value}</font></para>"
            ),
            styles['Body']
        )
        row.append(para)
        if idx % columns == 0:
            data.append(row)
            row = []
    if row:
        while len(row) < columns:
            row.append('')
        data.append(row)

    table = Table(data, colWidths=[col_width] * columns, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(BACKGROUND_COLOR)),
        ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    return table


def _owed_table(top_owed: List[Tuple[str, float]], styles, doc_width: float):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle

    if not top_owed:
        return Paragraph('No reimbursements owed this month.', styles['BodySmall'])

    display = top_owed[:5]
    extra = len(top_owed) - len(display)
    data = [
        [Paragraph('Employee', styles['BodySmall']), Paragraph('Amount', styles['BodySmall'])]
    ]
    for name, amount in display:
        data.append([
            Paragraph(name, styles['BodySmall']),
            Paragraph(formatCHF(amount), styles['BodySmall']),
        ])
    if extra > 0:
        data.append([
            Paragraph(f'+{extra} more', styles['Caption']),
            Paragraph('', styles['Caption'])
        ])

    table = Table(data, colWidths=[doc_width * 0.5, doc_width * 0.5], hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(BACKGROUND_COLOR)),
        ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return table


def _kpi_card_stack(values: List[Tuple[str, str]], doc_width: float, styles):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle

    table = Table([[Paragraph(
        (
            f"<para align=\"left\"><font name=\"{styles['HeadingSmall'].fontName}\" size=\"10\">{title}</font><br/>"
            f"<font name=\"{styles['HeadingLarge'].fontName}\" size=\"12\">{value}</font></para>"
        ),
        styles['Body'])] for title, value in values], colWidths=[doc_width], hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#FFFFFF')),
        ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    return table


def _heading_with_background(text: str, styles):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle

    heading = Paragraph(text, styles['HeadingMedium'])
    table = Table([[heading]], colWidths=['*'])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#E2E8F8')),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return table


def _totals_table(totals: Dict[str, Any], styles, doc_width: float):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle

    headers = [
        'Gross (CHF)',
        'Net (CHF)',
        'VAT (CHF)',
        'Company Card (CHF)',
        'Reimbursements Owed (CHF)',
        'Pending (CHF)',
    ]
    values = [
        formatCHF(totals.get('grossCHF')),
        formatCHF(totals.get('netCHF')),
        formatCHF(totals.get('vatCHF')),
        formatCHF(totals.get('companyCardSpentCHF')),
        formatCHF(totals.get('reimbursementsOwedCHF')),
        f"{formatCHF(totals.get('pendingTotalCHF'))}\n({totals.get('pendingTotalCount', 0)} items)",
    ]
    data = [
        [Paragraph(h, styles['BodySmall']) for h in headers],
        [Paragraph(v.replace("\n", "<br/>"), styles['BodySmall']) for v in values],
    ]
    table = Table(data, colWidths=[doc_width / len(headers)] * len(headers), hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(BACKGROUND_COLOR)),
        ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('ALIGN', (0, 1), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    return table


def _simple_table(headers: Iterable[str], rows: Iterable[Iterable[Any]], styles, doc_width: float, align_right_columns: Iterable[int] = (1,)):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle

    header_cells = [Paragraph(str(h), styles['BodySmall']) for h in headers]
    data = [header_cells]
    for row in rows:
        data.append([Paragraph(str(cell), styles['BodySmall']) for cell in row])

    column_count = len(header_cells)
    col_widths = [doc_width / column_count] * column_count

    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(BACKGROUND_COLOR)),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    for idx in align_right_columns:
        table.setStyle(TableStyle([('ALIGN', (idx, 1), (idx, -1), 'RIGHT')]))
    return table


def _bar_chart(data: List[Tuple[str, float]], title: str, width: float, height: float, base_font: str):
    from reportlab.graphics.charts.barcharts import HorizontalBarChart
    from reportlab.graphics.shapes import Drawing, String
    from reportlab.lib import colors

    if not data:
        return None

    labels = [item[0] for item in data]
    values = [float(item[1] or 0) for item in data]
    drawing = Drawing(width, height)
    chart = HorizontalBarChart()
    chart.x = 60
    chart.y = 20
    chart.height = height - 40
    chart.width = width - 120
    chart.data = [values]
    chart.categoryAxis.categoryNames = labels
    chart.barLabels.nudge = 8
    chart.barLabels.fontName = base_font
    chart.barLabels.fontSize = 8.5
    chart.categoryAxis.labels.fontName = base_font
    chart.categoryAxis.labels.fontSize = 8.5
    chart.valueAxis.labels.fontName = base_font
    chart.valueAxis.labels.fontSize = 8.5
    chart.valueAxis.strokeColor = colors.HexColor(LIGHT_BORDER)
    chart.bars[0].fillColor = colors.HexColor(ACCENT_COLOR)
    chart.valueAxis.valueMin = 0
    max_value = max(values)
    chart.valueAxis.valueMax = max_value * 1.15 if max_value else 1
    drawing.add(chart)
    drawing.add(String(0, height - 10, title, fontName=base_font, fontSize=10.5))
    return drawing


def _donut_chart(data: List[Tuple[str, float]], title: str, size: float, base_font: str):
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.shapes import Drawing, Circle, String
    from reportlab.lib import colors

    if not data:
        return None

    labels = [item[0] for item in data]
    values = [float(item[1] or 0) for item in data]
    if sum(values) == 0:
        return None

    drawing = Drawing(size, size)
    pie = Pie()
    pie.x = 10
    pie.y = 10
    pie.width = size - 20
    pie.height = size - 20
    pie.data = values
    pie.labels = labels
    pie.slices.strokeWidth = 0
    palette = [
        colors.HexColor('#2563EB'),
        colors.HexColor('#4F46E5'),
        colors.HexColor('#22C55E'),
        colors.HexColor('#F97316'),
        colors.HexColor('#06B6D4'),
        colors.HexColor('#EF4444'),
    ]
    for idx in range(len(values)):
        pie.slices[idx].fillColor = palette[idx % len(palette)]
    drawing.add(pie)
    drawing.add(Circle(size / 2, size / 2, size / 5, fillColor=colors.white, strokeColor=colors.white))
    drawing.add(String(0, size + 4, title, fontName=base_font, fontSize=10.5))
    return drawing


def _spark_bar_chart(data: List[Tuple[str, float]], width: float, height: float, base_font: str):
    from reportlab.graphics.charts.barcharts import HorizontalBarChart
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib import colors

    if not data:
        return None

    data = data[:8]
    labels = [item[0] for item in data]
    values = [float(item[1] or 0) for item in data]
    drawing = Drawing(width, height)
    chart = HorizontalBarChart()
    chart.x = 40
    chart.y = 10
    chart.height = height - 20
    chart.width = width - 60
    chart.data = [values]
    chart.categoryAxis.categoryNames = labels
    chart.categoryAxis.labels.fontName = base_font
    chart.categoryAxis.labels.fontSize = 8
    chart.valueAxis.labels.fontName = base_font
    chart.valueAxis.labels.fontSize = 8
    chart.valueAxis.valueMin = 0
    chart.valueAxis.visible = False
    chart.bars[0].fillColor = colors.HexColor('#10B981')
    drawing.add(chart)
    return drawing


class _NumberedCanvas:
    def __init__(self, base_canvas, footer_text: str, right_margin: float, bottom_margin: float, font_name: str):
        self._base_canvas = base_canvas
        self._footer_text = footer_text
        self._right_margin = right_margin
        self._bottom_margin = bottom_margin
        self._font_name = font_name
        self._saved_pages: List[Any] = []

    def showPage(self):
        self._saved_pages.append(dict(self._base_canvas.__dict__))
        self._base_canvas._startPage()

    def save(self):
        total = len(self._saved_pages)
        for state in self._saved_pages:
            self._base_canvas.__dict__.update(state)
            self._draw_footer(total)
            self._base_canvas.showPage()
        self._base_canvas.save()

    def __getattr__(self, name):
        return getattr(self._base_canvas, name)

    def _draw_footer(self, total_pages: int):
        self._base_canvas.saveState()
        self._base_canvas.setFont(self._font_name, 9)
        text = f"{self._footer_text} | page {self._base_canvas._pageNumber} of {total_pages}"
        x = self._base_canvas._pagesize[0] - self._right_margin
        y = self._bottom_margin / 2
        self._base_canvas.drawRightString(x, y, text)
        self._base_canvas.restoreState()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def render_report_pdf(report: Dict[str, Any]) -> bytes:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.lib import colors
        from reportlab.lib.utils import ImageReader
        from reportlab.platypus import (
            BaseDocTemplate,
            Frame,
            Image,
            PageTemplate,
            Paragraph,
            Spacer,
            Table,
            TableStyle,
            PageBreak,
        )
        from reportlab.pdfgen import canvas
    except Exception as exc:
        raise RuntimeError(
            'PDF generation requires reportlab. Install it: pip install reportlab'
        ) from exc

    base_font, bold_font = _register_fonts()
    styles = _load_styles(base_font, bold_font)

    aggregates = buildTables(report)
    period = str(report.get('period') or '')
    start_date, end_date = _period_bounds(period)
    generated_iso = datetime.utcnow().replace(microsecond=0).isoformat() + 'Z'

    metrics = {**aggregates, 'period': period}
    summary_text = summarizeForAI(metrics)

    buf = BytesIO()
    margin = 22 * mm
    doc = BaseDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')
    doc.addPageTemplates([PageTemplate(id='Pages', frames=[frame])])

    story: List[Any] = []

    # Page 1 -----------------------------------------------------------------
    logo_data = _logo_bytes()
    if logo_data:
        reader = ImageReader(BytesIO(logo_data))
        logo_width = 110
        iw, ih = reader.getSize()
        logo_height = logo_width * (ih / iw)
        logo = Image(BytesIO(logo_data), width=logo_width, height=logo_height)
    else:
        logo = Paragraph('Company Logo', styles['BodySmall'])

    title_para = Paragraph(
        (
            f"<para align=\"right\"><font name=\"{bold_font}\" size=\"16\">Monthly Expense Report — {period}</font><br/>"
            f"<font name=\"{base_font}\" size=\"10\">Generated {generated_iso}<br/>"
            f"Period {start_date.isoformat()} → {end_date.isoformat()}</font></para>"
        ),
        styles['Body']
    )
    header_table = Table([[logo, title_para]], colWidths=[doc.width * 0.4, doc.width * 0.6])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#EEF2FF')),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(header_table)
    story.append(Spacer(0, 10))

    totals = aggregates['totals']
    story.append(_metric_cards([
        ('Total Gross (CHF)', formatCHF(totals.get('grossCHF'))),
        ('Total Net (CHF)', formatCHF(totals.get('netCHF'))),
        ('Total VAT (CHF)', formatCHF(totals.get('vatCHF'))),
        ('Company Card Spent (CHF)', formatCHF(totals.get('companyCardSpentCHF'))),
    ], doc.width, styles))
    story.append(Spacer(0, 8))

    pending = aggregates['pending']
    reimb_total = totals.get('reimbursementsOwedCHF', 0.0)
    pending_total = totals.get('pendingTotalCHF', 0.0)
    in_progress = pending.get('inProgress', {})
    under_review = pending.get('underReview', {})
    pending_text = (
        f"{formatCHF(pending_total)}<br/><font size='9'>{in_progress.get('count', 0)} in-progress · "
        f"{formatCHF(in_progress.get('amount', 0))}<br/>{under_review.get('count', 0)} under review · "
        f"{formatCHF(under_review.get('amount', 0))}</font>"
    )
    story.append(
        _metric_cards([
            ('Reimbursements Owed Total (CHF)', formatCHF(reimb_total)),
            ('Pending Total (CHF)', pending_text),
        ], doc.width, styles, columns=2)
    )
    story.append(Spacer(0, 8))

    story.append(_heading_with_background('Who is owed what', styles))
    story.append(Spacer(0, 4))
    story.append(_owed_table(aggregates['topOwed'], styles, doc.width * 0.5))
    story.append(Spacer(0, 10))

    story.append(_heading_with_background('AI summary', styles))
    story.append(Spacer(0, 4))
    story.append(Paragraph(summary_text, styles['Body']))

    story.append(PageBreak())

    # Page 2 -----------------------------------------------------------------
    story.append(Paragraph('Numbers for accounting', styles['HeadingLarge']))
    story.append(_totals_table(totals, styles, doc.width))
    story.append(Spacer(0, 10))

    gross_total = totals.get('grossCHF') or 0.0
    story.append(_heading_with_background('By Category (CHF)', styles))
    category_rows = [
        (name, formatCHF(amount), percent(amount, gross_total))
        for name, amount in aggregates['byCategory']
    ]
    if category_rows:
        story.append(_simple_table(['Category', 'Gross', '% of total'], category_rows, styles, doc.width, align_right_columns=(1, 2)))
    else:
        story.append(Paragraph('No completed expenses recorded.', styles['BodySmall']))
    story.append(Spacer(0, 8))

    story.append(_heading_with_background('By Payment Method (CHF)', styles))
    payment_rows = [
        (name, formatCHF(amount), percent(amount, gross_total))
        for name, amount in aggregates['byPaymentMethod']
    ]
    if payment_rows:
        story.append(_simple_table(['Method', 'Gross', '% of total'], payment_rows, styles, doc.width, align_right_columns=(1, 2)))
    else:
        story.append(Paragraph('No completed expenses recorded.', styles['BodySmall']))
    story.append(Spacer(0, 8))

    story.append(_heading_with_background('Reimbursements by Employee (CHF)', styles))
    if aggregates['reimbursements']:
        reimb_rows = [
            (name, formatCHF(amount))
            for name, amount in aggregates['reimbursements']
        ]
        story.append(_simple_table(['Employee', 'Amount Owed'], reimb_rows, styles, doc.width, align_right_columns=(1,)))
    else:
        story.append(Paragraph('No reimbursements owed this month.', styles['BodySmall']))
    story.append(Spacer(0, 8))

    story.append(Paragraph(f"All values converted to CHF using {aggregates['ratePolicy']}. Originals available per row.", styles['Caption']))

    story.append(PageBreak())

    # Page 3+ ---------------------------------------------------------------
    story.append(Paragraph('Visual analysis', styles['HeadingLarge']))

    chart_data_category = sorted(aggregates['byCategory'], key=lambda item: item[1], reverse=True)
    chart_data_payment = sorted(aggregates['byPaymentMethod'], key=lambda item: item[1], reverse=True)
    chart_reimb = sorted(aggregates['reimbursements'], key=lambda item: item[1], reverse=True)

    bar = _bar_chart(chart_data_category, 'Gross by Category (CHF)', width=doc.width, height=200, base_font=base_font)
    donut = _donut_chart(chart_data_payment, 'Gross by Payment Method (CHF)', size=220, base_font=base_font)
    spark = _spark_bar_chart(chart_reimb, width=doc.width / 2, height=120, base_font=base_font)

    chart_row: List[Any] = []
    if bar:
        chart_row.append(bar)
    if donut:
        chart_row.append(donut)
    if chart_row:
        chart_table = Table([chart_row], colWidths=[doc.width / len(chart_row)] * len(chart_row))
        chart_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        story.append(chart_table)
        story.append(Spacer(0, 10))

    if spark:
        story.append(_heading_with_background('Reimbursements owed (top 8)', styles))
        story.append(spark)
        story.append(Spacer(0, 10))

    story.append(_heading_with_background('Pending overview', styles))
    pending_cards = [
        ('In-Progress', f"{in_progress.get('count', 0)} items<br/>{formatCHF(in_progress.get('amount', 0))}"),
        ('Under Review', f"{under_review.get('count', 0)} items<br/>{formatCHF(under_review.get('amount', 0))}"),
    ]
    story.append(_metric_cards(pending_cards, doc.width, styles, columns=2))
    story.append(Spacer(0, 10))

    story.append(Paragraph('Detailed rows', styles['HeadingMedium']))
    detail_rows = aggregates['rowsCHF']
    if not detail_rows:
        story.append(Paragraph('No rows for this period with the selected statuses.', styles['Body']))
    else:
        table_data = [[
            'Date', 'Payer', 'Category', 'Payment Method', 'Gross CHF', 'Net CHF', 'VAT CHF', 'Status', 'Original', 'Receipt'
        ]]
        for row in detail_rows:
            original = '-'
            if row['originalAmount']:
                original = f"{row['originalAmount']:.2f} {row['originalCurrency']}"
            link = row['receiptUrl']
            if link:
                link_cell = Paragraph(
                    f"<font color='{ACCENT_COLOR}'><link href='{link}'>View</link></font>",
                    styles['BodySmall']
                )
            else:
                link_cell = Paragraph('—', styles['BodySmall'])
            table_data.append([
                row.get('date', ''),
                row.get('payer', ''),
                row.get('category', ''),
                row.get('paymentMethod', ''),
                f"{row.get('amountCHF', 0.0):,.2f}".replace(',', "'"),
                f"{row.get('netCHF', 0.0):,.2f}".replace(',', "'"),
                f"{row.get('vatCHF', 0.0):,.2f}".replace(',', "'"),
                row.get('status', ''),
                original,
                link_cell,
            ])

        detail_table = Table(table_data, repeatRows=1, colWidths=[
            doc.width * 0.10,
            doc.width * 0.12,
            doc.width * 0.14,
            doc.width * 0.13,
            doc.width * 0.10,
            doc.width * 0.10,
            doc.width * 0.08,
            doc.width * 0.09,
            doc.width * 0.09,
            doc.width * 0.05,
        ])
        detail_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(BACKGROUND_COLOR)),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
            ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT_BORDER)),
            ('ALIGN', (4, 1), (6, -1), 'RIGHT'),
            ('ALIGN', (8, 1), (8, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 9.2),
        ]))
        for idx in range(1, len(table_data)):
            status_text = str(table_data[idx][7] or '').lower()
            color = STATUS_COLORS.get(status_text, STATUS_COLORS.get(status_text.replace('-', ' '), '#2563EB'))
            detail_table.setStyle(TableStyle([
                ('BACKGROUND', (7, idx), (7, idx), colors.HexColor(color)),
                ('TEXTCOLOR', (7, idx), (7, idx), colors.white),
            ]))
        story.append(detail_table)

    footer_text = 'In-House Expensify — Confidential'

    def _canvas_maker(*args, **kwargs):
        base = canvas.Canvas(*args, **kwargs)
        return _NumberedCanvas(base, footer_text, doc.rightMargin, doc.bottomMargin, base_font)

    doc.build(story, canvasmaker=_canvas_maker)
    return buf.getvalue()


def render_error_pdf(message: str) -> bytes:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.platypus import BaseDocTemplate, Paragraph, Frame, PageTemplate
    except Exception:
        return (message or 'Error').encode('utf-8')

    base_font, bold_font = _register_fonts()
    buf = BytesIO()
    doc = BaseDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')
    doc.addPageTemplates([PageTemplate(id='error', frames=[frame])])
    title = ParagraphStyle(name='Title', fontName=bold_font, fontSize=16, leading=20)
    body = ParagraphStyle(name='Body', fontName=base_font, fontSize=11, leading=14)
    story = [Paragraph('Report Generation Error', title), Paragraph(message or 'Unknown error', body)]
    doc.build(story)
    return buf.getvalue()
