'===========================================================
' APN KPI Annual Report 2026 - PowerPoint Macro Generator
' Creates 15 slides (template), tables, tree diagram, charts
' Data is placeholder (TBD) - ready to be filled from Excel later
'
' HOW TO USE:
' 1) Open PowerPoint -> ALT+F11 -> Insert > Module
' 2) Paste this whole code into the module
' 3) Run: Build_APN_KPI_Annual_2026_Template
'
' PATCH HISTORY:
' v1 - Initial release
' v2 - Fix Chr(8226) bullet error
' v3 - Fix compile error: remove AddChart2, replace xlLineMarkers/xlColumnClustered
'      with numeric constants (no Excel library reference needed)
'===========================================================
Option Explicit

'========================
' DESIGN TOKENS (APN Logo Palette)
'========================
Private Const APN_YELLOW As Long = &HF0F8
Private Const APN_GOLD   As Long = &HC0F8
Private Const APN_BLUE   As Long = &HD86830
Private Const APN_GREEN  As Long = &H408020
Private Const APN_MAROON As Long = &H2818A0
Private Const TEXT_BLACK As Long = &H111111
Private Const GRID_GRAY  As Long = &HD9D9D9

'========================
' CHART TYPE CONSTANTS (replaces xl* Excel constants)
' Numeric values from Excel ChartType enum - no Excel reference needed
'========================
Private Const CHART_LINE_MARKERS  As Long = 65   ' xlLineMarkers
Private Const CHART_COLUMN        As Long = 51   ' xlColumnClustered

'========================
' SLIDE SIZE
'========================
Private Const SLIDE_W As Single = 960
Private Const SLIDE_H As Single = 540

'========================
' LAYOUT CONSTANTS
'========================
Private Const M_LEFT   As Single = 40
Private Const M_TOP    As Single = 28
Private Const M_RIGHT  As Single = 40
Private Const M_BOTTOM As Single = 28
Private Const TITLE_H  As Single = 48

'========================
' ENTRYPOINT
'========================
Public Sub Build_APN_KPI_Annual_2026_Template()
    Dim pres As Presentation
    Set pres = ActivePresentation

    On Error Resume Next
    pres.PageSetup.SlideSize = ppSlideSizeOnScreen16x9
    On Error GoTo 0

    Dim i As Long
    For i = pres.Slides.Count To 1 Step -1
        pres.Slides(i).Delete
    Next i

    CreateSlide_01_Cover pres
    CreateSlide_02_Dashboard pres
    CreateSlide_03_TreeOrg pres
    CreateSlide_04_Bridging pres
    CreateSlide_05_SDM_Summary pres
    CreateSlide_06_SDM_Recruitment pres
    CreateSlide_07_SDM_DevelopmentHI pres
    CreateSlide_08_Training_Table pres
    CreateSlide_09_Training_LineChart pres
    CreateSlide_10_GA_Table pres
    CreateSlide_11_GA_ComboChart pres
    CreateSlide_12_K3LH_Table pres
    CreateSlide_13_Security_Table pres
    CreateSlide_14_Payroll_Split pres
    CreateSlide_15_ActionPlan pres

    MsgBox "Done: APN KPI Annual Report 2026 template (15 slides) created.", vbInformation
End Sub

'========================
' SLIDE BUILDERS
'========================

Private Sub CreateSlide_01_Cover(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddTitle sld, "LAPORAN PENCAPAIAN KPI TAHUNAN 2026"
    AddSubTitle sld, "DIREKTORAT SDM DAN UMUM"
    AddAccentBand sld, APN_GOLD, 0, SLIDE_H - 26, SLIDE_W, 10
    AddAccentBand sld, APN_BLUE, 0, SLIDE_H - 16, SLIDE_W, 6
    AddBox sld, SLIDE_W - 230, 70, 190, 190, "LOGO APN" & vbCrLf & "(insert image)", _
           TEXT_BLACK, 12, True, &HFFFFFF, GRID_GRAY
    AddSmallNote sld, "PT Agrinas Palma Nusantara", M_LEFT, SLIDE_H - 48
End Sub

Private Sub CreateSlide_02_Dashboard(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "RINGKASAN 39 KPI (DASHBOARD TOTAL)"
    Dim headers As Variant, rows As Variant
    headers = Array("Total KPI", "Achieved", "Not Achieved", "% Achievement")
    rows = Array(Array("39", "(diisi)", "(diisi)", "(diisi)"))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, 460, 130, headers, rows, APN_BLUE
    AddChartPlaceholder_Pie sld, M_LEFT + 500, M_TOP + TITLE_H + 10, 420, 260, "Komposisi Status KPI (Placeholder)"
    AddBulletBox sld, M_LEFT, M_TOP + TITLE_H + 160, 460, 170, _
        "HIGHLIGHT UTAMA (diisi setelah data final):", _
        Array("Area risiko tertinggi: (diisi)", "KPI merah dominan: (diisi)", "Catatan khusus Direksi: (diisi)")
End Sub

Private Sub CreateSlide_03_TreeOrg(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "STRUKTUR DIREKTORAT, DIVISI & SUBDIV (TREE DIAGRAM)"
    Dim root As String
    root = "Direktorat SDM & Umum"
    Dim branches As Variant
    branches = Array( _
        "Divisi Manajemen & Pengembangan SDM", _
        "Divisi Pusat Pelatihan", _
        "Divisi Umum & Manajemen Aset", _
        "Divisi K3LH & Pengamanan", _
        "Divisi Remunerasi & Penggajian" _
    )
    AddSimpleTree sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 360, root, branches
    AddSmallNote sld, "Catatan: Subdiv akan dijabarkan pada slide detail per divisi.", M_LEFT, SLIDE_H - 45
End Sub

Private Sub CreateSlide_04_Bridging(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "BRIDGING: CARA MEMBACA LAPORAN KPI"
    AddStatusLegend sld, M_LEFT, M_TOP + TITLE_H + 20
    AddBulletBox sld, M_LEFT, M_TOP + TITLE_H + 110, 520, 220, _
        "Definisi & Aturan:", _
        Array( _
            "Target: standar kinerja yang ditetapkan.", _
            "Actual: realisasi capaian pada periode pelaporan.", _
            "Status: Actual >= Target = ACH; Actual < Target = NOT; Data belum final = TBD." _
        )
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = Array(Array("1", "Contoh Parameter KPI", "(diisi)", "(diisi)", "? TBD"))
    AddAuditTable sld, M_LEFT + 560, M_TOP + TITLE_H + 110, 360, 120, headers, rows, APN_GOLD
    AddSmallNote sld, "Format tabel di seluruh slide konsisten (audit-style).", M_LEFT, SLIDE_H - 45
End Sub

Private Sub CreateSlide_05_SDM_Summary(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "DIVISI MANAJEMEN & PENGEMBANGAN SDM - RINGKASAN"
    Dim headers As Variant, rows As Variant
    headers = Array("Divisi", "Total KPI", "ACH", "NOT", "% Achievement")
    rows = Array(Array("Manajemen & Pengembangan SDM", "(diisi)", "(diisi)", "(diisi)", "(diisi)"))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 120, headers, rows, APN_BLUE
    AddBulletBox sld, M_LEFT, M_TOP + TITLE_H + 150, SLIDE_W - M_LEFT - M_RIGHT, 260, _
        "Top Issues (Placeholder):", _
        Array("KPI merah utama: (diisi)", "Area perbaikan: (diisi)", "Dukungan yang dibutuhkan: (diisi)")
End Sub

Private Sub CreateSlide_06_SDM_Recruitment(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "SDM - KINERJA REKRUTMEN & PEMENUHAN (DETAIL)"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Recruitment Fulfillment Rate", _
        "Fulfillment Lead Time", _
        "Internal Fill Rate", _
        "KPI Cascading Rate", _
        "Parameter tambahan (opsional)" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 320, headers, rows, APN_GOLD
    AddSmallNote sld, "Catatan: Target/Actual diisi dari Excel. Status default TBD.", M_LEFT, SLIDE_H - 45
End Sub

Private Sub CreateSlide_07_SDM_DevelopmentHI(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "SDM - PENGEMBANGAN ORGANISASI & HUBUNGAN INDUSTRIAL (DETAIL)"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Organization Effectiveness", _
        "Minimal Kandidat Suksesor", _
        "Jumlah Talent Siap (STAR)", _
        "Review STO", _
        "Turn Over Karyawan", _
        "Hari Kerja Tidak Efektif", _
        "Lembur > 56 Jam", _
        "Keluhan Karyawan", _
        "Akurasi Data Karyawan" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 380, headers, rows, APN_BLUE
End Sub

Private Sub CreateSlide_08_Training_Table(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "DIVISI PUSAT PELATIHAN - KINERJA PELATIHAN & PENGEMBANGAN"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Jam Training Karyawan", _
        "Kepuasan Peserta Training", _
        "Nilai Post Test Training", _
        "Parameter tambahan (opsional)" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 300, headers, rows, APN_GOLD
End Sub

Private Sub CreateSlide_09_Training_LineChart(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "PUSAT PELATIHAN - TREN KONSISTENSI PELATIHAN (JAN-DES)"
    Dim months As Variant, values As Variant
    months = Array("Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des")
    values = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    AddLineChart sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 340, _
                 "Jam Training per Karyawan (Placeholder)", months, values
    AddBulletBox sld, M_LEFT, M_TOP + TITLE_H + 365, SLIDE_W - M_LEFT - M_RIGHT, 120, _
        "Catatan:", _
        Array("Grafik diisi otomatis dari Excel (Jan-Des).", "Insight ditulis setelah data final tersedia.")
End Sub

Private Sub CreateSlide_10_GA_Table(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "DIVISI UMUM & MANAJEMEN ASET - KINERJA LAYANAN (DETAIL)"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Kecepatan Perbaikan", _
        "Resolution Rate", _
        "Kesiapan Kendaraan", _
        "Parameter layanan lainnya (opsional)" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 320, headers, rows, APN_BLUE
End Sub

Private Sub CreateSlide_11_GA_ComboChart(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "UMUM & ASET - ANALISA TREN PERBAIKAN ASET (COMBO CHART)"
    Dim months As Variant, aduan As Variant, speed As Variant
    months = Array("Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des")
    aduan  = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    speed  = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    AddComboChart_BarLine sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 360, _
                          "Total Aduan (Bar) vs Kecepatan Perbaikan (Line) - Placeholder", _
                          months, aduan, speed
    AddSmallNote sld, "Highlight bulan yang > target (rule diterapkan saat data final masuk).", M_LEFT, SLIDE_H - 45
End Sub

Private Sub CreateSlide_12_K3LH_Table(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "DIVISI K3LH - INDIKATOR KESELAMATAN KERJA (SHE)"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Zero Fatality", _
        "LTIFR", _
        "Zero Occupational Illness", _
        "Indikator tambahan (opsional)" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 300, headers, rows, APN_GOLD
    AddBox sld, M_LEFT, M_TOP + TITLE_H + 330, SLIDE_W - M_LEFT - M_RIGHT, 80, _
           "NOTE: Data Verification in Progress", TEXT_BLACK, 14, True, &HFFFFFF, APN_GOLD
End Sub

Private Sub CreateSlide_13_Security_Table(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "DIVISI PENGAMANAN - KINERJA PENGAMANAN & KEPATUHAN"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Kejadian Kehilangan", _
        "Pelanggaran Area", _
        "Patroli Rutin", _
        "Fungsi CCTV", _
        "Parameter tambahan (opsional)" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, 680, 360, headers, rows, APN_BLUE
    AddIconPlaceholder sld, M_LEFT + 710, M_TOP + TITLE_H + 40,  210, 100, "ICON CCTV"
    AddIconPlaceholder sld, M_LEFT + 710, M_TOP + TITLE_H + 160, 210, 100, "ICON SECURITY GUARD"
End Sub

Private Sub CreateSlide_14_Payroll_Split(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "DIVISI REMUNERASI & PENGGAJIAN - EFISIENSI & AKURASI PAYROLL"
    Dim headers As Variant, rows As Variant
    headers = Array("No", "Parameter", "Target", "Actual", "Status")
    rows = MakePlaceholderRows(Array( _
        "Akurasi Payroll (Ketepatan Waktu)", _
        "Akurasi Payroll (Ketepatan Hitung)", _
        "Employee Cost per Ton CPO (Target)", _
        "Parameter tambahan (opsional)" _
    ))
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, 440, 340, headers, rows, APN_GOLD
    Dim months As Variant, costs As Variant, tgt As Variant
    months = Array("Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des")
    costs  = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    tgt    = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    AddTwoSeriesLineChart sld, M_LEFT + 470, M_TOP + TITLE_H + 10, 450, 340, _
                          "Employee Cost per Ton CPO vs Target (Placeholder)", _
                          months, costs, tgt
End Sub

Private Sub CreateSlide_15_ActionPlan(ByVal pres As Presentation)
    Dim sld As Slide
    Set sld = AddBlankSlide(pres)
    AddSlideHeader sld, "PRIORITY ACTION PLAN 2027"
    Dim headers As Variant, rows As Variant
    headers = Array("Issue", "Root Cause", "Action", "Due Date", "Owner")
    rows = Array( _
        Array("(diisi)", "(diisi)", "(diisi)", "(diisi)", "(diisi)"), _
        Array("(diisi)", "(diisi)", "(diisi)", "(diisi)", "(diisi)"), _
        Array("(diisi)", "(diisi)", "(diisi)", "(diisi)", "(diisi)"), _
        Array("(diisi)", "(diisi)", "(diisi)", "(diisi)", "(diisi)"), _
        Array("(diisi)", "(diisi)", "(diisi)", "(diisi)", "(diisi)") _
    )
    AddAuditTable sld, M_LEFT, M_TOP + TITLE_H + 10, SLIDE_W - M_LEFT - M_RIGHT, 360, headers, rows, APN_BLUE
    AddSmallNote sld, "Action plan diisi setelah KPI merah tervalidasi. Prioritas: SDM, Umum, K3LH.", M_LEFT, SLIDE_H - 45
End Sub

'========================
' HELPERS: SLIDES
'========================

Private Function AddBlankSlide(ByVal pres As Presentation) As Slide
    Dim idx As Long
    idx = pres.Slides.Count + 1
    Set AddBlankSlide = pres.Slides.Add(idx, ppLayoutBlank)
    With AddBlankSlide.Background.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
End Function

Private Sub AddSlideHeader(ByVal sld As Slide, ByVal titleText As String)
    AddTitleAt sld, titleText, M_LEFT, M_TOP, SLIDE_W - M_LEFT - M_RIGHT, TITLE_H, 28
    AddAccentBand sld, APN_GOLD, M_LEFT, M_TOP + TITLE_H + 2, 240, 4
End Sub

Private Sub AddTitle(ByVal sld As Slide, ByVal titleText As String)
    AddTitleAt sld, titleText, M_LEFT, 70, SLIDE_W - M_LEFT - M_RIGHT, 80, 34
End Sub

Private Sub AddSubTitle(ByVal sld As Slide, ByVal subtitleText As String)
    AddTitleAt sld, subtitleText, M_LEFT, 155, SLIDE_W - M_LEFT - M_RIGHT, 40, 20
End Sub

Private Sub AddTitleAt(ByVal sld As Slide, ByVal txt As String, _
                       ByVal x As Single, ByVal y As Single, _
                       ByVal w As Single, ByVal h As Single, _
                       ByVal fontSize As Integer)
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    With shp.TextFrame2
        .TextRange.Text = txt
        With .TextRange.Font
            .Name = "Calibri"
            .Size = fontSize
            .Bold = msoTrue
            .Fill.ForeColor.RGB = TEXT_BLACK
        End With
        .VerticalAnchor = msoAnchorMiddle
    End With
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
End Sub

Private Sub AddSmallNote(ByVal sld As Slide, ByVal txt As String, _
                         ByVal x As Single, ByVal y As Single)
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, SLIDE_W - x - M_RIGHT, 20)
    With shp.TextFrame2
        .TextRange.Text = txt
        With .TextRange.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = msoFalse
            .Fill.ForeColor.RGB = &H666666
        End With
        .VerticalAnchor = msoAnchorMiddle
    End With
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
End Sub

Private Sub AddAccentBand(ByVal sld As Slide, ByVal colorRGB As Long, _
                          ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    shp.Line.Visible = msoFalse
    shp.Fill.ForeColor.RGB = colorRGB
    shp.Fill.Solid
End Sub

Private Sub AddBox(ByVal sld As Slide, _
                   ByVal x As Single, ByVal y As Single, _
                   ByVal w As Single, ByVal h As Single, _
                   ByVal txt As String, ByVal textColor As Long, _
                   ByVal fontSize As Integer, ByVal isBold As Boolean, _
                   ByVal fillColor As Long, ByVal borderColor As Long)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.Solid
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = borderColor
    shp.Line.Weight = 1
    With shp.TextFrame2
        .TextRange.Text = txt
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = "Calibri"
            .Size = fontSize
            .Bold = IIf(isBold, msoTrue, msoFalse)
            .Fill.ForeColor.RGB = textColor
        End With
    End With
End Sub

Private Sub AddIconPlaceholder(ByVal sld As Slide, _
                                ByVal x As Single, ByVal y As Single, _
                                ByVal w As Single, ByVal h As Single, _
                                ByVal lbl As String)
    AddBox sld, x, y, w, h, lbl & vbCrLf & "(insert icon)", TEXT_BLACK, 12, True, &HFFFFFF, GRID_GRAY
End Sub

Private Sub AddChartPlaceholder_Pie(ByVal sld As Slide, _
                                     ByVal x As Single, ByVal y As Single, _
                                     ByVal w As Single, ByVal h As Single, _
                                     ByVal lbl As String)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    shp.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shp.Fill.Solid
    shp.Line.ForeColor.RGB = GRID_GRAY
    shp.Line.Weight = 1
    shp.Line.DashStyle = msoLineDash
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 20
        .MarginRight = 20
        .MarginTop = 20
        .MarginBottom = 20
        With .TextRange
            .Text = lbl & vbCrLf & vbCrLf & _
                    "[Chart Pie/Donut akan ditambahkan" & vbCrLf & _
                    "setelah data final tersedia]"
            .Font.Size = 12
            .Font.Bold = msoFalse
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
        End With
    End With
End Sub

Private Sub AddBulletBox(ByVal sld As Slide, _
                         ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal h As Single, _
                         ByVal title As String, ByVal bullets As Variant)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Fill.ForeColor.RGB = &HFFFFFF
    shp.Fill.Solid
    shp.Line.ForeColor.RGB = GRID_GRAY
    shp.Line.Weight = 1
    Dim txt As String, i As Long
    txt = title & vbCrLf
    For i = LBound(bullets) To UBound(bullets)
        txt = txt & "• " & CStr(bullets(i)) & vbCrLf
    Next i
    With shp.TextFrame2
        .MarginLeft = 14
        .MarginRight = 12
        .MarginTop = 10
        .MarginBottom = 10
        .TextRange.Text = txt
        .VerticalAnchor = msoAnchorTop
        With .TextRange.Font
            .Name = "Calibri"
            .Size = 13
            .Fill.ForeColor.RGB = TEXT_BLACK
        End With
        .TextRange.Characters(1, Len(title)).Font.Bold = msoTrue
    End With
End Sub

'========================
' HELPERS: TABLES
'========================

Private Sub AddAuditTable(ByVal sld As Slide, _
                          ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single, _
                          ByVal headers As Variant, ByVal rows As Variant, _
                          ByVal headerColor As Long)
    Dim nCols As Long, nRows As Long
    nCols = UBound(headers) - LBound(headers) + 1
    nRows = UBound(rows) - LBound(rows) + 1
    Dim shp As Shape
    Set shp = sld.Shapes.AddTable(nRows + 1, nCols, x, y, w, h)
    Dim tbl As Table
    Set tbl = shp.Table
    StyleTableBorders tbl, GRID_GRAY, 0.75
    Dim c As Long
    For c = 1 To nCols
        SetCellText tbl.Cell(1, c), CStr(headers(LBound(headers) + c - 1)), 13, True, vbWhite
        tbl.Cell(1, c).Shape.Fill.ForeColor.RGB = headerColor
        tbl.Cell(1, c).Shape.Fill.Solid
        tbl.Cell(1, c).Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Next c
    Dim r As Long
    For r = 1 To nRows
        For c = 1 To nCols
            SetCellText tbl.Cell(r + 1, c), CStr(rows(LBound(rows) + r - 1)(c - 1)), 12, False, TEXT_BLACK
            tbl.Cell(r + 1, c).Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            tbl.Cell(r + 1, c).Shape.Fill.Solid
            tbl.Cell(r + 1, c).Shape.TextFrame2.MarginLeft = 8
            tbl.Cell(r + 1, c).Shape.TextFrame2.MarginRight = 6
            tbl.Cell(r + 1, c).Shape.TextFrame2.MarginTop = 4
            tbl.Cell(r + 1, c).Shape.TextFrame2.MarginBottom = 4
            tbl.Cell(r + 1, c).Shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Next c
        If (r Mod 2 = 0) Then
            For c = 1 To nCols
                tbl.Cell(r + 1, c).Shape.Fill.ForeColor.RGB = RGB(248, 248, 248)
                tbl.Cell(r + 1, c).Shape.Fill.Solid
            Next c
        End If
    Next r
    AutoAlignAuditColumns tbl, headers
End Sub

Private Sub StyleTableBorders(ByVal tbl As Table, ByVal borderColor As Long, ByVal wt As Single)
    Dim r As Long, c As Long
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Borders(ppBorderTop)
                .ForeColor.RGB = borderColor
                .Weight = wt
            End With
            With tbl.Cell(r, c).Borders(ppBorderBottom)
                .ForeColor.RGB = borderColor
                .Weight = wt
            End With
            With tbl.Cell(r, c).Borders(ppBorderLeft)
                .ForeColor.RGB = borderColor
                .Weight = wt
            End With
            With tbl.Cell(r, c).Borders(ppBorderRight)
                .ForeColor.RGB = borderColor
                .Weight = wt
            End With
        Next c
    Next r
End Sub

Private Sub SetCellText(ByVal cel As Cell, ByVal txt As String, _
                        ByVal sz As Integer, ByVal isBold As Boolean, _
                        ByVal color As Long)
    With cel.Shape.TextFrame2
        .TextRange.Text = txt
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange.Font
            .Name = "Calibri"
            .Size = sz
            .Bold = IIf(isBold, msoTrue, msoFalse)
            .Fill.ForeColor.RGB = color
        End With
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub

Private Sub AutoAlignAuditColumns(ByVal tbl As Table, ByVal headers As Variant)
    Dim c As Long, colName As String
    Dim align As MsoParagraphAlignment
    For c = 1 To tbl.Columns.Count
        colName = LCase$(CStr(headers(LBound(headers) + c - 1)))
        align = msoAlignLeft
        If InStr(colName, "no") > 0 Then align = msoAlignCenter
        If InStr(colName, "target") > 0 Or InStr(colName, "actual") > 0 Or InStr(colName, "%") > 0 Then align = msoAlignRight
        If InStr(colName, "status") > 0 Then align = msoAlignCenter
        Dim r As Long
        For r = 1 To tbl.Rows.Count
            tbl.Cell(r, c).Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = align
        Next r
    Next c
End Sub

Private Function MakePlaceholderRows(ByVal paramNames As Variant) As Variant
    Dim i As Long
    Dim rows() As Variant
    ReDim rows(0 To UBound(paramNames) - LBound(paramNames))
    For i = LBound(paramNames) To UBound(paramNames)
        rows(i - LBound(paramNames)) = Array( _
            CStr(i - LBound(paramNames) + 1), _
            CStr(paramNames(i)), _
            "(diisi dari Excel)", _
            "(diisi dari Excel)", _
            "? TBD" _
        )
    Next i
    MakePlaceholderRows = rows
End Function

'========================
' HELPERS: STATUS LEGEND
'========================

Private Sub AddStatusLegend(ByVal sld As Slide, ByVal x As Single, ByVal y As Single)
    AddStatusBadge sld, x, y, 180, 36, "ACH (Achieved)", APN_GREEN
    AddStatusBadge sld, x + 200, y, 200, 36, "NOT (Not Achieved)", RGB(200, 0, 0)
    AddStatusBadge sld, x + 420, y, 180, 36, "TBD (Verification)", APN_GOLD
End Sub

Private Sub AddStatusBadge(ByVal sld As Slide, _
                            ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single, _
                            ByVal txt As String, ByVal fillColor As Long)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.Solid
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .TextRange.Text = txt
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = "Calibri"
            .Size = 12
            .Bold = msoTrue
            .Fill.ForeColor.RGB = vbWhite
        End With
    End With
End Sub

'========================
' HELPERS: TREE DIAGRAM
'========================

Private Sub AddSimpleTree(ByVal sld As Slide, _
                          ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single, _
                          ByVal rootText As String, ByVal branches As Variant)
    Dim rootW As Single, rootH As Single
    rootW = 320: rootH = 52
    Dim rootX As Single, rootY As Single
    rootX = x + (w - rootW) / 2
    rootY = y
    Dim rootShp As Shape
    Set rootShp = sld.Shapes.AddShape(msoShapeRoundedRectangle, rootX, rootY, rootW, rootH)
    StyleTreeNode rootShp, rootText, APN_BLUE
    Dim n As Long
    n = UBound(branches) - LBound(branches) + 1
    Dim rowY As Single
    rowY = y + 120
    Dim nodeW As Single, nodeH As Single, gap As Single
    nodeW = 160: nodeH = 46
    gap = (w - (n * nodeW)) / (n + 1)
    If gap < 8 Then gap = 8
    Dim i As Long
    Dim nodeX As Single
    nodeX = x + gap
    Dim nodes() As Shape
    ReDim nodes(1 To n)
    For i = 1 To n
        Set nodes(i) = sld.Shapes.AddShape(msoShapeRoundedRectangle, nodeX, rowY, nodeW, nodeH)
        StyleTreeNode nodes(i), CStr(branches(LBound(branches) + i - 1)), APN_GOLD
        nodeX = nodeX + nodeW + gap
    Next i
    Dim midRootX As Single, midRootY As Single
    midRootX = rootX + rootW / 2
    midRootY = rootY + rootH
    For i = 1 To n
        Dim midNodeX As Single, midNodeY As Single
        midNodeX = nodes(i).Left + nodes(i).Width / 2
        midNodeY = nodes(i).Top
        AddConnectorLine sld, midRootX, midRootY, midNodeX, midNodeY, GRID_GRAY
    Next i
    Dim noteY As Single
    noteY = rowY + 70
    AddBox sld, x, noteY, w, h - (noteY - y), _
           "SUBDIV (placeholder):" & vbCrLf & _
           "- Setiap divisi memiliki subdiv sesuai struktur internal." & vbCrLf & _
           "- Detail subdiv akan ditampilkan pada slide breakdown.", _
           TEXT_BLACK, 11, False, RGB(255, 255, 255), GRID_GRAY
End Sub

Private Sub StyleTreeNode(ByVal shp As Shape, ByVal txt As String, ByVal fillColor As Long)
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.Solid
    shp.Line.ForeColor.RGB = GRID_GRAY
    shp.Line.Weight = 1
    With shp.TextFrame2
        .MarginLeft = 8: .MarginRight = 8
        .TextRange.Text = txt
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = msoTrue
            .Fill.ForeColor.RGB = vbWhite
        End With
    End With
End Sub

Private Sub AddConnectorLine(ByVal sld As Slide, _
                              ByVal x1 As Single, ByVal y1 As Single, _
                              ByVal x2 As Single, ByVal y2 As Single, _
                              ByVal color As Long)
    Dim ln As Shape
    Set ln = sld.Shapes.AddLine(x1, y1, x2, y2)
    ln.Line.ForeColor.RGB = color
    ln.Line.Weight = 1.25
End Sub

'========================
' HELPERS: CHARTS
' PATCH v3: Removed AddChart2 (PPT 2016+ only).
'           Replaced xlLineMarkers/xlColumnClustered with numeric constants.
'           No Excel library reference needed.
'========================

Private Sub AddLineChart(ByVal sld As Slide, _
                         ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal h As Single, _
                         ByVal title As String, ByVal categories As Variant, _
                         ByVal values As Variant)
    Dim shp As Shape
    On Error Resume Next
    Set shp = sld.Shapes.AddChart(CHART_LINE_MARKERS, x, y, w, h)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    With shp.Chart
        .HasTitle = True
        .ChartTitle.Text = title
        On Error Resume Next
        .ChartArea.Format.Fill.Visible = msoFalse
        On Error GoTo 0
        SetChartData_OneSeries .ChartData, "Jam Training", categories, values
    End With
End Sub

Private Sub AddTwoSeriesLineChart(ByVal sld As Slide, _
                                  ByVal x As Single, ByVal y As Single, _
                                  ByVal w As Single, ByVal h As Single, _
                                  ByVal title As String, ByVal categories As Variant, _
                                  ByVal series1 As Variant, ByVal series2 As Variant)
    Dim shp As Shape
    On Error Resume Next
    Set shp = sld.Shapes.AddChart(CHART_LINE_MARKERS, x, y, w, h)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    With shp.Chart
        .HasTitle = True
        .ChartTitle.Text = title
        On Error Resume Next
        .ChartArea.Format.Fill.Visible = msoFalse
        On Error GoTo 0
        SetChartData_TwoSeries .ChartData, categories, "Actual", series1, "Target", series2
    End With
End Sub

Private Sub AddComboChart_BarLine(ByVal sld As Slide, _
                                  ByVal x As Single, ByVal y As Single, _
                                  ByVal w As Single, ByVal h As Single, _
                                  ByVal title As String, ByVal categories As Variant, _
                                  ByVal barSeries As Variant, ByVal lineSeries As Variant)
    Dim shp As Shape
    On Error Resume Next
    Set shp = sld.Shapes.AddChart(CHART_COLUMN, x, y, w, h)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    With shp.Chart
        .HasTitle = True
        .ChartTitle.Text = title
        On Error Resume Next
        .ChartArea.Format.Fill.Visible = msoFalse
        SetChartData_TwoSeries .ChartData, categories, "Total Aduan", barSeries, "Kecepatan Perbaikan", lineSeries
        ' Change series 2 to line type (65 = xlLineMarkers)
        .SeriesCollection(2).ChartType = CHART_LINE_MARKERS
        On Error GoTo 0
    End With
End Sub

'========================
' CHART DATA WRITERS
'========================

Private Sub SetChartData_OneSeries(ByVal cd As ChartData, ByVal seriesName As String, _
                                   ByVal categories As Variant, ByVal values As Variant)
    On Error Resume Next
    cd.Activate
    Dim wb As Object, ws As Object
    Set wb = cd.Workbook
    Set ws = wb.Worksheets(1)
    ws.Cells.Clear
    ws.Range("A1").Value = "Month"
    ws.Range("B1").Value = seriesName
    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        ws.Cells(i + 2, 1).Value = categories(i)
        ws.Cells(i + 2, 2).Value = values(i)
    Next i
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub SetChartData_TwoSeries(ByVal cd As ChartData, ByVal categories As Variant, _
                                   ByVal s1Name As String, ByVal s1Values As Variant, _
                                   ByVal s2Name As String, ByVal s2Values As Variant)
    On Error Resume Next
    cd.Activate
    Dim wb As Object, ws As Object
    Set wb = cd.Workbook
    Set ws = wb.Worksheets(1)
    ws.Cells.Clear
    ws.Range("A1").Value = "Month"
    ws.Range("B1").Value = s1Name
    ws.Range("C1").Value = s2Name
    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        ws.Cells(i + 2, 1).Value = categories(i)
        ws.Cells(i + 2, 2).Value = s1Values(i)
        ws.Cells(i + 2, 3).Value = s2Values(i)
    Next i
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub