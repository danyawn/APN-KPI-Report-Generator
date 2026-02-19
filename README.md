# APN KPI Annual Report Generator

**VBA PowerPoint Macro** — Auto-generate a 15-slide KPI Annual Report template for **PT Agrinas Palma Nusantara 2026** (HR & General Affairs Directorate).

---

## 📋 Overview

This VBA macro creates a fully structured PowerPoint presentation with:

- **15 slides** covering:
  - Cover slide
  - KPI dashboard summary (39 total KPIs)
  - Organizational tree diagram
  - Bridging slide (how to read the report)
  - 5 division-specific slides (SDM, Training, GA, K3LH, Security, Payroll)
  - Detailed KPI tables (audit-style)
  - Line charts, combo charts (bar + line)
  - Action plan for 2027

- **Design tokens** matching APN logo palette (yellow, gold, blue, green, maroon)
- **Audit-style tables** with alternating row colors and auto-aligned columns
- **Chart placeholders** ready for data population from Excel
- **Tree diagram** for organizational structure visualization

---

## 🚀 How to Use

1. **Open PowerPoint** (any blank presentation)
2. Press **ALT + F11** to open VBA Editor
3. Go to **Insert > Module**
4. **Copy-paste** the entire code from `APN_KPI_Report_Generator.bas` into the module window
5. Press **F5** or run macro: `Build_APN_KPI_Annual_2026_Template`
6. Done! 15 slides will be generated automatically.

---

## 🔧 Features

### Design System
- **APN Color Palette**:
  - Yellow: `#F8F000`
  - Gold: `#F8C000`
  - Blue: `#3068D8`
  - Green: `#208040`
  - Maroon: `#A01828`
- **16:9 Wide Layout** (960×540px)
- **Calibri font** throughout
- **Consistent spacing** (40px margins, 28px title height)

### Slide Types
1. **Cover** — Title + subtitle + logo placeholder + accent bands
2. **Dashboard Table** — Summary of 39 KPIs (Total/Achieved/Not Achieved/% Achievement) + pie chart placeholder
3. **Tree Diagram** — Directorate → 5 Divisions (with connector lines)
4. **Bridging Slide** — Legend (ACH/NOT/TBD status badges) + definition bullet box + example table
5. **Division Summary Slides** — Each division has:
   - Summary table (total KPIs per division)
   - Bullet box (top issues placeholder)
6. **Detail KPI Tables** — Audit-style tables with:
   - Header row (colored background: blue/gold)
   - Data rows (alternating white/light gray)
   - Auto-aligned columns (center for No/Status, right for Target/Actual/%)
7. **Line Charts** — Monthly trend data (Jan-Dec)
8. **Combo Chart** — Bar (Total Aduan) + Line (Kecepatan Perbaikan)
9. **Action Plan Table** — 5 rows: Issue / Root Cause / Action / Due Date / Owner

### Helpers
- **AddAuditTable** — Creates professional tables with borders and formatting
- **AddBulletBox** — Rounded rectangle with title + bullet points (hardcoded `•` for compatibility)
- **AddLineChart / AddTwoSeriesLineChart / AddComboChart_BarLine** — Chart generation with ChartData workbook
- **AddSimpleTree** — Organizational structure with root + branches + connector lines
- **AddStatusLegend** — Colored badges (ACH/NOT/TBD)
- **AddChartPlaceholder_Pie** — Dashed-border rectangle for future pie chart insertion

---

## 🐛 Troubleshooting

### "Invalid procedure call or argument" Error

**Cause**: Some PowerPoint versions have issues with `Chr(8226)` (bullet character).

**Fix**: This version uses **hardcoded Unicode bullet** (`"• "`) in `AddBulletBox` function instead of `Chr(8226)` for maximum compatibility.

### Charts Not Appearing

**Cause**: Chart data is currently **placeholder zeros** (all values = 0).

**Solution**: 
- Charts will display once you populate data from Excel (Phase 2: Excel integration).
- Current version shows chart placeholders with labels.

### Slide Layout Issues

**Cause**: PowerPoint theme/master slide conflicts.

**Solution**: 
- Macro uses `ppLayoutBlank` for all slides (bypasses theme).
- All styling is done programmatically (fills, fonts, borders).

---

## 📊 Data Population (Next Phase)

Currently, all KPI data shows:
- `"(diisi dari Excel)"` — Target/Actual fields
- `"? TBD"` — Status field (to be determined)

**Phase 2 Plan** (Excel integration):
1. Read KPI data from Excel workbook (39 parameters × 5 divisions)
2. Calculate status automatically:
   - `ACH` if Actual >= Target
   - `NOT` if Actual < Target
3. Populate chart data arrays (monthly trends)
4. Apply conditional formatting:
   - Green fill for ACH rows
   - Red fill for NOT rows
   - Yellow fill for TBD rows

---

## 📁 File Structure

```
APN-KPI-Report-Generator/
├── APN_KPI_Report_Generator.bas    # Main VBA code (copy-paste to PowerPoint)
├── README.md                        # This file
└── LICENSE                          # MIT License
```

---

## 🤝 Contributing

Contributions welcome! To improve:
- Add more chart types (scatter, pie, donut)
- Integrate direct Excel data pull
- Add animation effects
- Enhance tree diagram (multi-level subdiv)

---

## 📄 License

MIT License — Free to use, modify, and distribute.

---

## 👤 Author

**Yan Danu Tirta** ([@danyawn](https://github.com/danyawn))  
Software Engineering • UPN Veteran Yogyakarta  
Website: [yan-danu.vercel.app](https://yan-danu.vercel.app/)

---

## 🎯 Use Case

Perfect for:
- Corporate annual KPI reporting
- HR & General Affairs performance dashboards
- Board/Director presentations
- Audit-ready documentation
- Multi-division performance tracking

---

**Last Updated**: February 19, 2026  
**PowerPoint Compatibility**: 2016, 2019, 2021, Microsoft 365