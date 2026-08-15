"""Content and writer for the 'Summary charts' tab.

Every chart points at cells on the data tabs - nothing is copied or recomputed here, so
a chart can never drift from the numbers it claims to show. Rebuilding the workbook
rebuilds the charts from the same source rows.

Colours are a fixed categorical order, assigned per series and never cycled: a series
keeps its colour across every chart it appears in, so NSW is the same blue on the daily,
monthly, quarterly and annual electricity charts. The palette was validated for
colour-vision deficiency separation (worst adjacent pair dE 9.1 protan, normal-vision
19.6). Three of the five slots sit under 3:1 contrast on white, which is acceptable here
because identity never rests on colour alone - every chart carries a legend, and the
underlying table is one tab away.
"""
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (CharacterProperties, Paragraph, ParagraphProperties,
                                   RichTextProperties)
from openpyxl.styles import Font

# Validated categorical palette, in fixed slot order.
PALETTE = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7", "E34948"]
GRID = "E8E8E6"
AXIS = "8C8C8C"
HDR_ROW = 4          # data tabs put their header on row 4
FIRST_DATA_ROW = 5
LINE_W = 19050       # ~1.5pt in EMU

TITLE_FONT = Font(bold=True, size=13, color="1F3864")
SUB_FONT = Font(italic=True, size=9, color="595959")

# The petrol charts are pinned to a common scale so that wholesale and retail, and each
# frequency, can be read against one another without rescaling by eye. The band is set
# wide enough to contain every charted petrol value - the 2020 TGP trough at 80.8 c/L and
# the April 2026 peak at 325.9 - so nothing is drawn off the edge of a plot. Widen it if a
# future refresh moves outside; the assertion in check_ylim() will say so.
# Electricity and gas stay auto-scaled - their spikes are orders of magnitude, not
# percentages, and a fixed range would either flatten them or clip them off.
PETROL_YLIM = (75, 350, 50)   # min, max, gridline spacing

# (data tab stem, chart title, y-axis title, [(column header, legend name), ...], y-range)
SERIES = [
    ("Electricity", "Electricity spot price (DWA)", "$/MWh",
     [("NSW1_price_dwa", "NSW"), ("QLD1_price_dwa", "QLD"), ("SA1_price_dwa", "SA"),
      ("TAS1_price_dwa", "TAS"), ("VIC1_price_dwa", "VIC")], None),
    ("Gas", "Gas wholesale price", "$/GJ",
     [("VIC_DWGM_6am", "VIC DWGM"), ("SYD_exante", "Sydney STTM"),
      ("ADL_exante", "Adelaide STTM"), ("BRI_exante", "Brisbane STTM")], None),
    ("Petrol TGP", "Petrol wholesale TGP, national", "cents/litre",
     [("TGP_petrol_national", "Petrol"), ("TGP_diesel_national", "Diesel")], PETROL_YLIM),
    ("Petrol retail", "Retail ULP pump price", "cents/litre",
     [("NSW_ULP", "NSW"), ("QLD_ULP", "QLD"), ("WA_ULP", "WA")], PETROL_YLIM),
]

# Row selectors. None is the whole tab; ("window", n) is the trailing n years measured
# from that series' own last date; ("year", y) is calendar year y. The fourth item is an
# axis override for that chart alone; without one the series' own y-range applies.
STANDARD_VARIANTS = [
    ("daily", "daily, last 2 years", ("window", 2), None),
    ("daily", "daily, last 5 years", ("window", 5), None),
    ("monthly", "monthly", None, None),
    ("quarterly", "quarterly", None, None),
    ("annual", "annual", None, None),
]
# Petrol also gets a chart per recent calendar year, each on a tighter axis of its own so
# a single year's movement is legible rather than flattened into the full 75-350 band.
# 2026 has to reach 350 rather than a tighter ceiling: TGP diesel peaked at 325.9 c/L on
# 9 April 2026, and anything lower crops that spike off the top of the plot.
YLIM_2026 = (125, 350, 25)
YLIM_2025 = (125, 200, 25)
PETROL_VARIANTS = STANDARD_VARIANTS[:2] + [
    ("daily", "daily, 2026", ("year", 2026), YLIM_2026),
    ("daily", "daily, 2025", ("year", 2025), YLIM_2025),
] + STANDARD_VARIANTS[2:]

# One electricity/gas pair per state, each on its own two-scale chart. Tasmania has no
# gas hub - it is not in the STTM and has no DWGM - so it has no chart here; TAS spot
# prices are on the electricity tabs. Each pair starts where its gas series starts,
# which differs by hub.
STATE_PAIRS = [
    ("New South Wales", "NSW1_price_dwa", "NSW spot", "SYD_exante", "Sydney STTM"),
    ("Victoria", "VIC1_price_dwa", "VIC spot", "VIC_DWGM_6am", "VIC DWGM"),
    ("Queensland", "QLD1_price_dwa", "QLD spot", "BRI_exante", "Brisbane STTM"),
    ("South Australia", "SA1_price_dwa", "SA spot", "ADL_exante", "Adelaide STTM"),
]


def col_index(ws, header):
    """1-based column of a header on a data tab, or None if the tab lacks it."""
    for j in range(1, ws.max_column + 1):
        if ws.cell(row=HDR_ROW, column=j).value == header:
            return j
    return None


def search_rows(ws, ok):
    """First and last row whose date column A satisfies ok(), or None if none do."""
    rows = [i for i in range(FIRST_DATA_ROW, ws.max_row + 1)
            if ok(ws.cell(row=i, column=1).value)]
    return (rows[0], rows[-1]) if rows else None


def select_rows(ws, selector):
    """Resolve a variant's row selector to (first_row, last_row)."""
    if selector is None:
        return FIRST_DATA_ROW, ws.max_row
    kind, n = selector
    if kind == "window":
        cutoff = ws.cell(row=ws.max_row, column=1).value.toordinal() - round(365.25 * n)
        return search_rows(ws, lambda d: d.toordinal() >= cutoff)
    if kind == "year":
        return search_rows(ws, lambda d: d.year == n)
    raise ValueError(f"unknown row selector {selector!r}")


def year_label(ws, first_row, last_row, year):
    """Say so when a calendar-year chart is only part of a year."""
    last = ws.cell(row=last_row, column=1).value
    if (last.month, last.day) == (12, 31):
        return f"daily, {year}"
    return f"daily, {year} (to {last.strftime('%-d %b')})"


def row_of(ws, key):
    """Row holding a given period key, matching on the tab's own column A values."""
    for i in range(FIRST_DATA_ROW, ws.max_row + 1):
        if ws.cell(row=i, column=1).value == key:
            return i
    return None


def style_series(s, colour):
    s.graphicalProperties = GraphicalProperties(
        ln=LineProperties(solidFill=colour, w=LINE_W))
    s.smooth = False
    s.marker.symbol = "none"


def recede_axes(chart, n_points):
    gl = ChartLines()
    gl.spPr = GraphicalProperties(ln=LineProperties(solidFill=GRID, w=9525))
    chart.y_axis.majorGridlines = gl
    chart.x_axis.majorGridlines = None
    for ax in (chart.x_axis, chart.y_axis):
        ax.delete = False
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill=AXIS, w=9525))
    skip = max(1, n_points // 8)
    chart.x_axis.tickLblSkip = skip
    chart.x_axis.tickMarkSkip = skip
    # Period labels are wide enough to collide when laid flat, so lay them on the slant.
    small = CharacterProperties(sz=800)
    chart.x_axis.txPr = RichText(
        bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=small), endParaRPr=small)])



def check_ylim(ws, cols, first_row, last_row, ylim, chart_title, deliberate):
    """A fixed axis hides anything outside it, so never let that pass unremarked.

    The shared petrol band is meant to contain everything, so anything outside it is a
    mistake and stops the build. An axis chosen for one chart is an editorial decision -
    it is allowed to crop, but every build says what it crops, so the cost stays visible.
    """
    lo, hi = ylim[0], ylim[1]
    outside = []
    for header, name in cols:
        j = col_index(ws, header)
        if j is None:
            continue
        for i in range(first_row, last_row + 1):
            v = ws.cell(row=i, column=j).value
            if v is not None and not lo <= v <= hi:
                outside.append((ws.cell(row=i, column=1).value, header, v))
    if not outside:
        return
    if not deliberate:
        d, header, v = outside[0]
        raise SystemExit(
            f"ERROR: {ws.title} {header} = {v} on {d} falls outside the fixed axis range "
            f"{lo}-{hi}, so the chart would draw it off the plot. Widen PETROL_YLIM in "
            f"build_charts.py.")
    days = sorted({d for d, _, _ in outside})
    worst = max(outside, key=lambda t: abs(t[2] - (hi if t[2] > hi else lo)))
    print(f"  note: '{chart_title}' axis {lo}-{hi} crops {len(days)} day(s), "
          f"{days[0]:%Y-%m-%d} to {days[-1]:%Y-%m-%d}; furthest is {worst[1]} "
          f"{worst[2]} on {worst[0]:%Y-%m-%d}", flush=True)


def line_chart(ws, title, y_title, cols, first_row, last_row, ylim=None,
               width=15.5, height=8.2):
    chart = LineChart()
    chart.title = title
    chart.y_axis.title = y_title
    chart.height, chart.width = height, width
    if ylim:
        chart.y_axis.scaling.min, chart.y_axis.scaling.max, chart.y_axis.majorUnit = ylim
    for k, (header, name) in enumerate(cols):
        j = col_index(ws, header)
        if j is None:
            continue
        s = Series(Reference(ws, min_col=j, min_row=first_row, max_row=last_row),
                   title=name)
        style_series(s, PALETTE[k % len(PALETTE)])
        chart.append(s)
    chart.set_categories(Reference(ws, min_col=1, min_row=first_row, max_row=last_row))
    recede_axes(chart, last_row - first_row + 1)
    if len(chart.series) > 1:
        chart.legend.position = "b"
        chart.legend.overlay = False
    else:
        chart.legend = None
    return chart


def first_populated_row(ws, col):
    """First row where a column actually has a number - hubs start at different dates."""
    for i in range(FIRST_DATA_ROW, ws.max_row + 1):
        if ws.cell(row=i, column=col).value is not None:
            return i
    return None


def combined_chart(wb, state, elec_col, elec_name, gas_col, gas_name):
    """One state's electricity and gas, quarterly, with gas on a secondary axis.

    Both series are from the same state, so the two lines describe one market rather than
    two unrelated ones. Each chart starts at the first quarter its gas hub published -
    2007-Q1 for the Victorian DWGM, 2010-Q3 for Sydney and Adelaide, 2011-Q4 for Brisbane
    - because charting from 1998 would leave most of the plot with only one line on it.

    The two prices are on genuinely different scales ($/MWh against $/GJ), so read the
    shapes and their timing against each other, not the crossings - where the lines sit
    relative to one another is an artefact of the two axis ranges, not a fact about the
    market.
    """
    elec, gas = wb["Electricity quarterly"], wb["Gas quarterly"]
    g0 = first_populated_row(gas, col_index(gas, gas_col))
    g1 = gas.max_row
    e0, e1 = row_of(elec, gas.cell(row=g0, column=1).value), elec.max_row
    if e1 - e0 != g1 - g0:
        raise SystemExit(f"ERROR: {state} quarterly rows do not line up between the "
                         f"electricity and gas tabs ({e1 - e0 + 1} vs {g1 - g0 + 1}); "
                         f"the two lines would be drawn against the wrong quarters.")

    c1 = LineChart()
    c1.title = f"Electricity and gas — quarterly, {state} (two scales)"
    c1.y_axis.title = "Electricity $/MWh"
    c1.y_axis.axId = 100
    s = Series(Reference(elec, min_col=col_index(elec, elec_col), min_row=e0, max_row=e1),
               title=f"{elec_name} ($/MWh, left)")
    style_series(s, PALETTE[0])
    c1.append(s)
    c1.set_categories(Reference(elec, min_col=1, min_row=e0, max_row=e1))
    recede_axes(c1, e1 - e0 + 1)
    c1.height, c1.width = 8.2, 15.5

    c2 = LineChart()
    c2.y_axis.axId = 200
    c2.y_axis.title = "Gas $/GJ"
    c2.y_axis.majorGridlines = None
    c2.y_axis.delete = False
    c2.y_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=AXIS, w=9525))
    s = Series(Reference(gas, min_col=col_index(gas, gas_col), min_row=g0, max_row=g1),
               title=f"{gas_name} ($/GJ, right)")
    style_series(s, PALETTE[1])
    c2.append(s)

    # Draws the gas axis on the right. This must be set on the secondary axis: setting
    # it on the primary instead moves the electricity axis across, leaving each title on
    # the opposite side from its own scale.
    c2.y_axis.crosses = "max"
    c1 += c2
    c1.legend.position = "b"
    c1.legend.overlay = False
    return c1


def iter_slots(row=4, pitch=18):
    """Anchors for a two-column grid of charts, top to bottom."""
    while True:
        yield f"B{row}"
        yield f"L{row}"
        row += pitch


def write_charts(wb):
    ws = wb.create_sheet("Summary charts", 1)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Summary charts"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Every series at each frequency, drawn live from the data tabs — the "
                "numbers behind any chart are on the matching tab. Daily is shown over "
                "the last two years and the last five, and for petrol over 2026 and 2025 "
                "separately; monthly, quarterly and annual show full history. A series "
                "keeps the same colour across all of its charts. The petrol charts are "
                "pinned to a fixed 75–350 c/L scale so they can be read against each "
                "other. The four two-scale charts pair each state's electricity with its "
                "own gas hub.")
    ws["A2"].font = SUB_FONT

    slots = iter_slots()
    for pair in STATE_PAIRS:
        ws.add_chart(combined_chart(wb, *pair), next(slots))

    for stem, title, y_title, cols, ylim in SERIES:
        variants = PETROL_VARIANTS if ylim else STANDARD_VARIANTS
        for freq, label, selector, ylim_override in variants:
            name = f"{stem} {freq}"
            if name not in wb.sheetnames:
                continue
            data = wb[name]
            rows = select_rows(data, selector)
            if rows is None:            # a year the series does not reach
                continue
            first, last = rows
            if selector and selector[0] == "year":
                label = year_label(data, first, last, selector[1])
            axis = ylim_override or ylim
            chart_title = f"{title} — {label}"
            if axis:
                check_ylim(data, cols, first, last, axis, chart_title,
                           deliberate=ylim_override is not None)
            chart = line_chart(data, chart_title, y_title, cols, first, last, ylim=axis)
            ws.add_chart(chart, next(slots))

    ws.column_dimensions["A"].width = 2
    return ws
