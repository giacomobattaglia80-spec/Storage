#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Add a third 'normalised' (2023+2024 aggregated) cost bar to the P&L slide."""
f = "/home/user/Storage/Hennecke/pitch/build_pptx_editable.py"
src = open(f).read()

NEW = r'''# =====================================================================
# SLIDE 2a2 — P&L COST STRUCTURE 2023 · 2024 · NORMALISED
# =====================================================================
s = slide()
header(s, "P&L recap — cost structure (2023 · 2024 · normalised)",
       "Summing both years nets out the project-cycle WIP swing → a normalised run-rate", "3 / 7")
cats = ["EBITDA", "Materials", "Personnel", "Services", "Lease & rentals", "Other op. costs"]
colA = {"EBITDA":"84BD00","Materials":"F2A100","Personnel":"D9531E","Services":"3C7DA6","Lease & rentals":"8C5A2B","Other op. costs":"BFBFBF"}
txtA = {"EBITDA":INK,"Materials":INK,"Personnel":WHITE,"Services":WHITE,"Lease & rentals":WHITE,"Other op. costs":INK}
d23 = {"EBITDA":7.0,"Materials":46.0,"Personnel":23.1,"Services":19.9,"Lease & rentals":3.8,"Other op. costs":0.3}
d24 = {"EBITDA":3.8,"Materials":38.4,"Personnel":30.3,"Services":22.8,"Lease & rentals":4.2,"Other op. costs":0.5}
dN  = {"EBITDA":5.6,"Materials":42.7,"Personnel":26.2,"Services":21.2,"Lease & rentals":4.0,"Other op. costs":0.4}
TOPB, SCALE, BW = 165, 6.0, 120
def draw_bar(x, data):
    y = TOPB; bounds = [y]
    for c in cats:
        h = data[c]*SCALE
        rect(s, x, y, BW, h, fill=colA[c])
        if h >= 17:
            para(tb(s, x, y, BW, h, "m"), f"{data[c]:.1f}%", size=(12 if h >= 26 else 9), color=txtA[c], bold=True, align="c", first=True)
        y += h; bounds.append(y)
    return bounds
BX1, BX2, BX3 = 430, 640, 850
b1 = draw_bar(BX1, d23); b2 = draw_bar(BX2, d24); b3 = draw_bar(BX3, dN)
for xa, ba, xb, bb in [(BX1, b1, BX2, b2), (BX2, b2, BX3, b3)]:
    for i in range(len(ba)):
        cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, PX(xa+BW), PX(ba[i]), PX(xb), PX(bb[i]))
        cn.line.color.rgb = RGBColor.from_string("D9DEE6"); cn.line.width = Pt(0.75); cn.shadow.inherit = False
rect(s, 802, 162, 1.6, 620, fill="C9D2DD")  # separator before normalised bar
bot = max(b1[-1], b2[-1], b3[-1])
para(tb(s, BX1-15, bot+8, BW+30, 22), "2023", size=15, color=BLUE, bold=True, align="c", first=True)
para(tb(s, BX2-15, bot+8, BW+30, 22), "2024", size=15, color=BLUE, bold=True, align="c", first=True)
para(tb(s, BX3-25, bot+8, BW+50, 22), "2-yr", size=15, color=BLUE, bold=True, align="c", first=True)
para(tb(s, BX3-45, bot+30, BW+90, 20), "(normalised)", size=11, color=GRAY, align="c", first=True)
para(tb(s, 300, TOPB-2, 120, 22), "% of value", size=11, color=MUTED, align="r", first=True)
para(tb(s, 300, TOPB+16, 120, 22), "of production", size=11, color=MUTED, align="r", first=True)
# recap table
rect(s, 1010, 190, 540, 300, fill=LIGHT, line=BORDER)
para(tb(s, 1030, 200, 220, 26), "Category", size=12, color=GRAY, bold=True, first=True)
para(tb(s, 1255, 200, 90, 26), "2023", size=12, color=GRAY, bold=True, align="c", first=True)
para(tb(s, 1355, 200, 90, 26), "2024", size=12, color=GRAY, bold=True, align="c", first=True)
para(tb(s, 1455, 200, 90, 26), "2-yr", size=12, color=BLUE, bold=True, align="c", first=True)
rect(s, 1030, 230, 510, 1, fill=BORDER)
ry = 244
for c in cats:
    para(tb(s, 1030, ry, 230, 28, "m"), [("■ ", False, colA[c], 12), (c, True, INK, 12)], first=True)
    para(tb(s, 1255, ry, 90, 28, "m"), f"{d23[c]:.1f}%", size=12, color=TXT, align="c", first=True)
    para(tb(s, 1355, ry, 90, 28, "m"), f"{d24[c]:.1f}%", size=12, color=TXT, align="c", first=True)
    para(tb(s, 1455, ry, 90, 28, "m"), f"{dN[c]:.1f}%", size=12, color=BLUE, bold=True, align="c", first=True)
    ry += 36
# callout
rect(s, 1010, 505, 540, 150, fill=BLUEBG, line=BLUELN)
tfc = tb(s, 1032, 515, 496, 132)
para(tfc, [("Normalising the cycle: ", True, BLUE, 14), ("aggregating 2023+2024 nets out the WIP swing, giving a run-rate EBITDA of ", False, TXT, 14), ("~5.6% of value of production", True, TXT, 14), (" (~5.5% of revenue) — a fairer valuation base than either single year.", False, TXT, 14)], first=True, after=6)
para(tfc, "Single years are distorted by project timing: 7.0% (2023, WIP build-up) vs 3.8% (2024, WIP unwind).", size=12, color=GRAY, italic=True)
para(tb(s, 50, 860, 1330, 30), "2-yr = 2023+2024 aggregated (value of production €92.7m), neutralising the project-cycle WIP swing. % of value of production; may not sum to 100 due to rounding. Source: Coface reclassified P&L.", size=9.5, color=MUTED, first=True)

'''

start = src.index("# SLIDE 2a2 ")
banner = src.rindex("# ====", 0, start)
end = src.index("# =====================================================================\n# SLIDE 2b — CUSTOMER MAPPING")
src = src[:banner] + NEW + src[end:]
open(f, "w").write(src)
print("Added normalised (2-yr) bar to P&L slide")
