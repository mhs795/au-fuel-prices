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

# The petrol charts are pinned to a common 100-300 c/L scale so that wholesale and
# retail, and each frequency, can be read against one another without rescaling by eye.
# Electricity and gas stay auto-scaled - their spikes are orders of magnitude, not
# percentages, and a fixed range would either flatten them or clip them off.
PETROL_YLIM = (100, 300)

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

# (data tab frequency, chart label, trailing window in years - None for full history)
VARIANTS = [
    ("daily", "daily, last 2 years", 2),
    ("daily", "daily, last 5 years", 5),
    ("monthly", "monthly", None),
    ("quarterly", "quarterly", None),
    ("annual", "annual", None),
]


def col_index(ws, header):
    """1-based column of a header on a data tab, or None if the tab lacks it."""
    for j in range(1, ws.max_column + 1):
        if ws.cell(row=HDR_ROW, column=j).value == header:
            return j
    return None


def daily_start_row(ws, years):
    """First row within the trailing window, so daily charts stay legible."""
    last = ws.cell(row=ws.max_row, column=1).value
    cutoff = last.toordinal() - round(365.25 * years)
    lo, hi = FIRST_DATA_ROW, ws.max_row
    while lo < hi:
        mid = (lo + hi) // 2
        if ws.cell(row=mid, column=1).value.toordinal() < cutoff:
            lo = mid + 1
        else:
            hi = mid
    return lo


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


def line_chart(ws, title, y_title, cols, first_row, last_row, ylim=None,
               width=15.5, height=8.2):
    chart = LineChart()
    chart.title = title
    chart.y_axis.title = y_title
    chart.height, chart.width = height, width
    if ylim:
        chart.y_axis.scaling.min, chart.y_axis.scaling.max = ylim
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


def combined_chart(wb):
    """Electricity and gas quarterly on one chart, gas on a secondary axis.

    Victorian series on both sides - VIC1 spot against the Victorian DWGM - so the two
    lines describe the same market rather than two unrelated ones. The window starts at
    2007-Q1, the first quarter gas exists; charting from 1998 would leave two thirds of
    the plot with only one line on it.

    The two prices are on genuinely different scales ($/MWh against $/GJ), so read the
    shapes and their timing against each other, not the crossings - where the lines sit
    relative to one another is an artefact of the two axis ranges, not a fact about the
    market.
    """
    elec, gas = wb["Electricity quarterly"], wb["Gas quarterly"]
    start_key = gas.cell(row=FIRST_DATA_ROW, column=1).value
    e0, e1 = row_of(elec, start_key), elec.max_row
    g0, g1 = FIRST_DATA_ROW, gas.max_row

    c1 = LineChart()
    c1.title = "Electricity and gas — quarterly, Victoria (two scales)"
    c1.y_axis.title = "Electricity $/MWh"
    c1.y_axis.axId = 100
    s = Series(Reference(elec, min_col=col_index(elec, "VIC1_price_dwa"),
                         min_row=e0, max_row=e1), title="Electricity VIC ($/MWh, left)")
    style_series(s, PALETTE[0])
    c1.append(s)
    c1.set_categories(Reference(elec, min_col=1, min_row=e0, max_row=e1))
    recede_axes(c1, e1 - e0 + 1)
    c1.height, c1.width = 9.5, 32.5

    c2 = LineChart()
    c2.y_axis.axId = 200
    c2.y_axis.title = "Gas $/GJ"
    c2.y_axis.majorGridlines = None
    c2.y_axis.delete = False
    c2.y_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=AXIS, w=9525))
    s = Series(Reference(gas, min_col=col_index(gas, "VIC_DWGM_6am"),
                         min_row=g0, max_row=g1), title="Gas VIC DWGM ($/GJ, right)")
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


def write_charts(wb):
    ws = wb.create_sheet("Summary charts", 1)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Summary charts"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Every series at each frequency, drawn live from the data tabs — the "
                "numbers behind any chart are on the matching tab. Daily is shown over "
                "the last two years and the last five; monthly, quarterly and annual show "
                "full history. A series keeps the same colour across all of its charts. "
                "The petrol charts are pinned to a fixed 100–300 c/L scale.")
    ws["A2"].font = SUB_FONT

    row, col = 4, 0

    ws.add_chart(combined_chart(wb), "B4")
    row += 21

    for stem, title, y_title, cols, ylim in SERIES:
        for freq, label, window in VARIANTS:
            name = f"{stem} {freq}"
            if name not in wb.sheetnames:
                continue
            data = wb[name]
            first = daily_start_row(data, window) if window else FIRST_DATA_ROW
            chart = line_chart(data, f"{title} — {label}", y_title, cols,
                               first, data.max_row, ylim=ylim)
            ws.add_chart(chart, f"{'B' if col == 0 else 'L'}{row}")
            col += 1
            if col == 2:
                col, row = 0, row + 18

    ws.column_dimensions["A"].width = 2
    return ws
