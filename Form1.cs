using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

namespace DataParser
{
    public partial class Form1 : Form
    {
        private string? _excelPath;
        private Point _dragStart;
        private bool _dragging;
        private static readonly string SettingsFile = Path.Combine(AppContext.BaseDirectory, "settings.json");
        private static readonly int ROW_COUNT = 13;
        private static readonly string[] DefaultKeywords = new[]
        {
            "ID",
            "Networker Name, networker",
            "1., Full Name, name",
            "2., Gender",
            "3., Birthday, date of birth, dob",
            "4., Contact, phone, mobile",
            "5., Occupation, job",
            "6., City of Residence, city, residence",
            "7., Church Name, church",
            "8., Denomination",
            "9., Church Position, position",
            "10., Married/Single, marital status",
            "11., Facebook address, facebook, fb"
        };

        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            this.Icon = CreateAutoInputIcon();
            EnableFormDrag();
            AddCloseButton();
            InitTable();
            LoadSettings();
        }

        private static Icon CreateAutoInputIcon()
        {
            int sz = 64;
            using var bmp = new Bitmap(sz, sz);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.Transparent);
            using var bgBrush = new SolidBrush(Color.FromArgb(22, 27, 34));
            using var bgPath = new System.Drawing.Drawing2D.GraphicsPath();
            int r = 12;
            bgPath.AddArc(0, 0, r, r, 180, 90); bgPath.AddArc(sz - r, 0, r, r, 270, 90);
            bgPath.AddArc(sz - r, sz - r, r, r, 0, 90); bgPath.AddArc(0, sz - r, r, r, 90, 90);
            bgPath.CloseFigure(); g.FillPath(bgBrush, bgPath);
            using var docPen = new Pen(Color.FromArgb(88, 166, 255), 2.2f);
            g.DrawRectangle(docPen, 12, 8, 22, 28);
            using var linePen = new Pen(Color.FromArgb(139, 148, 158), 1.5f);
            g.DrawLine(linePen, 16, 15, 30, 15); g.DrawLine(linePen, 16, 20, 28, 20); g.DrawLine(linePen, 16, 25, 26, 25);
            using var arrowPen = new Pen(Color.FromArgb(63, 185, 80), 3f);
            arrowPen.EndCap = System.Drawing.Drawing2D.LineCap.Round; arrowPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
            g.DrawLine(arrowPen, 44, 14, 44, 40); g.DrawLine(arrowPen, 37, 33, 44, 42); g.DrawLine(arrowPen, 51, 33, 44, 42);
            using var basePen = new Pen(Color.FromArgb(63, 185, 80), 2f);
            g.DrawLine(basePen, 36, 46, 52, 46);
            using var numFont = new Font("Consolas", 10f, FontStyle.Bold);
            using var numBrush = new SolidBrush(Color.FromArgb(88, 166, 255));
            g.DrawString("0", numFont, numBrush, 22, 40);
            return Icon.FromHandle(bmp.GetHicon());
        }

        private void InitTable()
        {
            for (int i = 0; i < DefaultKeywords.Length && i < _txtKeywords.Length; i++)
                _txtKeywords[i].Text = DefaultKeywords[i];
        }

        private void EnableFormDrag()
        {
            foreach (Control c in new Control[] { panelTop, lblInput, panelMid, lblResult, panelBottom })
            {
                c.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; } };
                c.MouseMove += (s, e) => { if (_dragging) { var p = PointToScreen(e.Location); SetDesktopLocation(p.X - _dragStart.X, p.Y - _dragStart.Y); } };
                c.MouseUp += (s, e) => _dragging = false;
            }
        }

        private void LayoutTable()
        {
            int w = panelTable.ClientSize.Width;
            int colW = 60, uniW = 50, rowH = 24;
            int titleW = (int)(w * 0.42);
            int valW = w - titleW - uniW - colW;
            int rowCount = _txtKeywords.Length;
            while (_hLines.Length <= rowCount)
            {
                var nl = new Panel { BackColor = Color.FromArgb(48, 54, 61), Height = 1 };
                panelTable.Controls.Add(nl);
                Array.Resize(ref _hLines, _hLines.Length + 1);
                _hLines[_hLines.Length - 1] = nl;
            }
            int y = 0;
            for (int i = 0; i < rowCount; i++)
            {
                _hLines[i].SetBounds(0, y, w, 1); _hLines[i].Visible = true; y += 1;
                _txtKeywords[i].SetBounds(4, y + 2, titleW - 8, rowH - 4);
                _txtValues[i].SetBounds(titleW + 4, y + 2, valW - 8, rowH - 4);
                _rdoUnique[i].SetBounds(titleW + valW + (uniW - 16) / 2, y + (rowH - 16) / 2, 16, 16);
                _txtColNums[i].SetBounds(titleW + valW + uniW + 4, y + 2, colW - 8, rowH - 4);
                y += rowH;
            }
            _hLines[rowCount].SetBounds(0, y, w, 1); _hLines[rowCount].Visible = true;
            for (int i = rowCount + 1; i < _hLines.Length; i++) _hLines[i].Visible = false;
            _vLine1.SetBounds(titleW, 0, 1, y + 1);
            _vLine2.SetBounds(titleW + valW, 0, 1, y + 1);
            _vLine3.SetBounds(titleW + valW + uniW, 0, 1, y + 1);
            _vLine1.BringToFront(); _vLine2.BringToFront(); _vLine3.BringToFront();
            panelTable.AutoScrollPosition = new Point(0, 0);
            panelTable.AutoScrollMinSize = new Size(0, y + 2);

            // 헤더도 동일한 w 기준으로 레이아웃 (스크롤바 포함된 동일 너비)
            int hdrH = panelHeader.Height;
            hdrTitle.SetBounds(0, 0, titleW, hdrH);
            hdrValue.SetBounds(titleW, 0, valW, hdrH);
            hdrUnique.SetBounds(titleW + valW, 0, uniW, hdrH);
            hdrCol.SetBounds(titleW + valW + uniW, 0, colW, hdrH);
        }

        private void AddCloseButton()
        {
            var b = new Label { Text = "\u2715", ForeColor = Color.FromArgb(139,148,158), BackColor = Color.Transparent,
                Font = new Font("Segoe UI",9F,FontStyle.Bold), Size = new Size(28,22),
                TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            b.Click += (s, e) => Close();
            b.MouseEnter += (s, e) => b.ForeColor = Color.FromArgb(248,81,73);
            b.MouseLeave += (s, e) => b.ForeColor = Color.FromArgb(139,148,158);
            this.Controls.Add(b); b.BringToFront();
            this.Resize += (s, e) => b.Location = new Point(this.ClientSize.Width - 32, 4);
            b.Location = new Point(this.ClientSize.Width - 32, 4);
        }

        private void SaveSettings()
        {
            try {
                var lines = new List<string> { "EXCEL=" + (_excelPath ?? ""), "ROWS=" + _txtKeywords.Length };
                // UNIQUE 선택 인덱스 저장
                int uniqueIdx = -1;
                for (int i = 0; i < _rdoUnique.Length; i++)
                    if (_rdoUnique[i].Checked) { uniqueIdx = i; break; }
                lines.Add("UNIQUE=" + uniqueIdx);
                for (int i = 0; i < _txtKeywords.Length; i++)
                { lines.Add($"KW{i}=" + _txtKeywords[i].Text.Trim()); lines.Add($"COL{i}=" + _txtColNums[i].Text.Trim()); }
                File.WriteAllLines(SettingsFile, lines);
            } catch { }
        }

        private void LoadSettings()
        {
            try {
                if (!File.Exists(SettingsFile)) return;
                var saved = new Dictionary<string, string>(); string? savedExcel = null; int savedRows = ROW_COUNT; int savedUnique = -1;
                foreach (string line in File.ReadAllLines(SettingsFile))
                { int eq = line.IndexOf('='); if (eq < 0) continue; string key = line[..eq], val = line[(eq+1)..];
                  if (key == "EXCEL") savedExcel = val;
                  else if (key == "ROWS" && int.TryParse(val, out int rc)) savedRows = rc;
                  else if (key == "UNIQUE" && int.TryParse(val, out int ui)) savedUnique = ui;
                  else saved[key] = val; }
                if (!string.IsNullOrEmpty(savedExcel) && File.Exists(savedExcel))
                { _excelPath = savedExcel; lblExcelPath.Text = Path.GetFileName(_excelPath); lblExcelPath.ForeColor = Color.FromArgb(88,166,255); LoadSheetNames(); }
                while (_txtKeywords.Length < savedRows) AddTableRow();
                for (int i = 0; i < _txtKeywords.Length; i++)
                { if (saved.TryGetValue($"KW{i}", out string? kw) && !string.IsNullOrEmpty(kw)) _txtKeywords[i].Text = kw;
                  if (saved.TryGetValue($"COL{i}", out string? col)) _txtColNums[i].Text = col; }
                // UNIQUE 복원
                if (savedUnique >= 0 && savedUnique < _rdoUnique.Length)
                {
                    for (int i = 0; i < _rdoUnique.Length; i++) _rdoUnique[i].Checked = false;
                    _rdoUnique[savedUnique].Checked = true;
                }
            } catch { }
        }

        private void BtnProcess_Click(object? sender, EventArgs e)
        {
            string raw = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(raw)) { MessageBox.Show("데이터를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            ParseAndDisplay(raw);
        }

        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Title = "엑셀 파일 선택", Filter = "Excel Files|*.xlsx;*.xls", RestoreDirectory = true };
            if (dlg.ShowDialog() == DialogResult.OK)
            { _excelPath = dlg.FileName; lblExcelPath.Text = Path.GetFileName(_excelPath); lblExcelPath.ForeColor = Color.FromArgb(88,166,255); LoadSheetNames(); SaveSettings(); }
        }

        private void LoadSheetNames()
        {
            cboSheet.Items.Clear();
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath)) return;
            try {
                using var pkg = new ExcelPackage(new FileInfo(_excelPath));
                int autoIdx = -1;
                for (int i = 0; i < pkg.Workbook.Worksheets.Count; i++)
                { string n = pkg.Workbook.Worksheets[i].Name; cboSheet.Items.Add(n);
                  if (n.Equals("Lahore2", StringComparison.OrdinalIgnoreCase)) autoIdx = i; }
                if (autoIdx >= 0) cboSheet.SelectedIndex = autoIdx; else if (cboSheet.Items.Count > 0) cboSheet.SelectedIndex = 0;
            } catch { }
        }

        private void BtnInsert_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            { MessageBox.Show("먼저 엑셀 파일을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (cboSheet.SelectedIndex < 0)
            { MessageBox.Show("엑셀 시트를 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // 키워드가 비어있는 행 확인
            var emptyRows = new List<int>();
            for (int i = 0; i < _txtKeywords.Length; i++)
                if (string.IsNullOrWhiteSpace(_txtKeywords[i].Text)) emptyRows.Add(i);

            if (emptyRows.Count > 0)
            {
                var result = MessageBox.Show("키워드 값이 없는 항목이 있습니다.\n해당 행을 삭제하시겠습니까?",
                    "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    RemoveRows(emptyRows);
                    SaveSettings();
                }
                ResetScroll();
            }

            string selectedSheet = cboSheet.SelectedItem?.ToString() ?? "";
            var entries = new List<(int col, string value)>();
            for (int i = 0; i < _txtKeywords.Length; i++)
            { string val = _txtValues[i].Text.Trim(); string numStr = _txtColNums[i].Text.Trim();
              if (!string.IsNullOrEmpty(numStr) && int.TryParse(numStr, out int colIdx) && colIdx > 0) entries.Add((colIdx, CleanSpecialChars(val))); }
            if (entries.Count == 0) { MessageBox.Show("컬럼 번호(COL No.)가 지정된 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (!entries.Any(x => !string.IsNullOrEmpty(x.value))) { MessageBox.Show("입력할 데이터가 없습니다.\n먼저 데이터를 입력하고 1차 가공을 해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // UNIQUE 선택 확인
            int uniqueIdx = -1;
            for (int i = 0; i < _rdoUnique.Length; i++)
                if (_rdoUnique[i].Checked) { uniqueIdx = i; break; }

            try {
                byte[] fileBytes = File.ReadAllBytes(_excelPath);
                using var stream = new MemoryStream(fileBytes);
                using var pkg = new ExcelPackage(stream);
                var ws = pkg.Workbook.Worksheets[selectedSheet];
                if (ws == null) { MessageBox.Show($"시트 '{selectedSheet}'를 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                int totalRows = ws.Dimension?.End.Row ?? 0; int totalCols = ws.Dimension?.End.Column ?? 0;

                // UNIQUE가 선택된 경우: 해당 값으로 기존 행 검색
                if (uniqueIdx >= 0)
                {
                    string uniqueValue = CleanSpecialChars(_txtValues[uniqueIdx].Text.Trim());
                    string uniqueColStr = _txtColNums[uniqueIdx].Text.Trim();
                    if (string.IsNullOrEmpty(uniqueValue)) { MessageBox.Show("UNIQUE로 선택된 항목의 VALUE가 비어있습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                    if (!int.TryParse(uniqueColStr, out int uniqueCol) || uniqueCol <= 0) { MessageBox.Show("UNIQUE로 선택된 항목의 COL No.가 지정되지 않았습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                    // 엑셀에서 해당 컬럼에서 동일한 값 검색
                    int foundRow = -1;
                    for (int row = 1; row <= totalRows; row++)
                    {
                        var cellVal = ws.Cells[row, uniqueCol].Value;
                        if (cellVal != null && cellVal.ToString()!.Trim().Equals(uniqueValue, StringComparison.OrdinalIgnoreCase))
                        { foundRow = row; break; }
                    }

                    if (foundRow > 0)
                    {
                        var dlgResult = MessageBox.Show(
                            $"지정된 엑셀시트 {foundRow}행에 유일 값으로 지정된 항목과 동일한 항목이 있어 추가하지 않고 값을 수정합니다.",
                            "알림", MessageBoxButtons.OKCancel, MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1, 0);

                        if (dlgResult == DialogResult.OK)
                        {
                            int updatedCount = 0;
                            var updatedItems = new List<string>();
                            for (int i = 0; i < _txtKeywords.Length; i++)
                            {
                                if (i == uniqueIdx) continue;
                                string val = _txtValues[i].Text.Trim();
                                string numStr = _txtColNums[i].Text.Trim();
                                if (string.IsNullOrEmpty(val)) continue;
                                if (!int.TryParse(numStr, out int col) || col <= 0) continue;
                                string cleanVal = CleanSpecialChars(val);
                                ws.Cells[foundRow, col].Value = cleanVal;
                                updatedItems.Add($"COL {col}: {cleanVal}");
                                updatedCount++;
                            }
                            File.WriteAllBytes(_excelPath, pkg.GetAsByteArray());
                            string itemList = updatedCount > 0 ? string.Join("\n", updatedItems) : "(없음)";
                            MessageBox.Show($"시트 '{selectedSheet}'의 {foundRow}행 업데이트 완료 ({updatedCount}개)\n\n{itemList}", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ResetAfterInsert();
                        }
                        else
                        {
                            var addResult = MessageBox.Show("기존 데이터를 무시하고 새 행에 추가하시겠습니까?",
                                "신규 추가", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (addResult == DialogResult.Yes)
                            {
                                int lastDataRow = 0;
                                for (int row = 1; row <= totalRows; row++)
                                    for (int c = 1; c <= totalCols; c++)
                                    { var v = ws.Cells[row, c].Value; if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) { lastDataRow = row; break; } }
                                int newRow = lastDataRow + 1;
                                for (int i = 0; i < _txtKeywords.Length; i++)
                                { string val = _txtValues[i].Text.Trim(); string numStr = _txtColNums[i].Text.Trim();
                                  if (!string.IsNullOrEmpty(numStr) && int.TryParse(numStr, out int col) && col > 0) ws.Cells[newRow, col].Value = CleanSpecialChars(val); }
                                File.WriteAllBytes(_excelPath, pkg.GetAsByteArray());
                                MessageBox.Show($"시트 '{selectedSheet}'의 {newRow}행에 새로 입력했습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                ResetAfterInsert();
                            }
                        }
                    }
                    else
                    {
                        int lastDataRow = 0;
                        for (int row = 1; row <= totalRows; row++)
                            for (int c = 1; c <= totalCols; c++)
                            { var v = ws.Cells[row, c].Value; if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) { lastDataRow = row; break; } }
                        int newRow = lastDataRow + 1;
                        for (int i = 0; i < _txtKeywords.Length; i++)
                        { string val = _txtValues[i].Text.Trim(); string numStr = _txtColNums[i].Text.Trim();
                          if (!string.IsNullOrEmpty(numStr) && int.TryParse(numStr, out int col) && col > 0) ws.Cells[newRow, col].Value = CleanSpecialChars(val); }
                        File.WriteAllBytes(_excelPath, pkg.GetAsByteArray());
                        MessageBox.Show($"시트 '{selectedSheet}'의 {newRow}행에 새로 입력했습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ResetAfterInsert();
                    }
                }
                else
                {
                    int lastDataRow = 0;
                    for (int row = 1; row <= totalRows; row++)
                        for (int c = 1; c <= totalCols; c++)
                        { var v = ws.Cells[row, c].Value; if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) { lastDataRow = row; break; } }
                    int newRow = lastDataRow + 1;
                    for (int i = 0; i < _txtKeywords.Length; i++)
                    { string val = _txtValues[i].Text.Trim(); string numStr = _txtColNums[i].Text.Trim();
                      if (!string.IsNullOrEmpty(numStr) && int.TryParse(numStr, out int col) && col > 0) ws.Cells[newRow, col].Value = CleanSpecialChars(val); }
                    File.WriteAllBytes(_excelPath, pkg.GetAsByteArray());
                    MessageBox.Show($"시트 '{selectedSheet}'의 {newRow}행에 데이터를 입력했습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ResetAfterInsert();
                }
            } catch (Exception ex) {
                string msg = ex.Message ?? "";
                string innerMsg = ex.InnerException?.Message ?? "";
                string allMsg = msg + " " + innerMsg;
                if (ex is IOException
                    || allMsg.Contains("being used by another process", StringComparison.OrdinalIgnoreCase)
                    || allMsg.Contains("locked", StringComparison.OrdinalIgnoreCase)
                    || allMsg.Contains("Error saving file", StringComparison.OrdinalIgnoreCase)
                    || allMsg.Contains("denied", StringComparison.OrdinalIgnoreCase))
                    MessageBox.Show("엑셀 파일이 열려 있어 수정할 수 없습니다.\n파일을 닫고 다시 시도해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else MessageBox.Show($"엑셀 저장 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetAfterInsert()
        { txtInput.Text = ""; for (int i = 0; i < _txtKeywords.Length; i++) _txtValues[i].Text = ""; ResetScroll(); }

        private void ParseAndDisplay(string rawText)
        {
            rawText = CleanFacebookAdd(rawText);
            string[] inputLines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < _txtKeywords.Length; i++)
            {
                string kwText = _txtKeywords[i].Text.Trim();
                if (string.IsNullOrEmpty(kwText)) { _txtValues[i].Text = ""; continue; }
                string[] keywords = kwText.Split(',').Select(k => k.Trim()).Where(k => k.Length > 0).ToArray();
                var textKws = new List<string>(); var numPats = new List<int>();
                foreach (string kw in keywords)
                { var m = Regex.Match(kw, @"^(\d+)\.$"); if (m.Success) numPats.Add(int.Parse(m.Groups[1].Value)); else textKws.Add(kw); }
                string? found = null;
                foreach (string kw in textKws) { found = SearchByTextKeyword(inputLines, kw); if (found != null) break; }
                if (found == null) foreach (int num in numPats) { found = SearchByNumberPattern(inputLines, num); if (found != null) break; }
                _txtValues[i].Text = found ?? "";
            }
        }

        private static string? SearchByTextKeyword(string[] lines, string keyword)
        {
            string kwLow = keyword.ToLower();
            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim(); if (string.IsNullOrEmpty(line)) continue;
                string stripped = Regex.Replace(line, @"^\d+\.\s*", "");
                int ci = stripped.IndexOf(':'); if (ci < 0) continue;
                string keyPart = Regex.Replace(stripped[..ci].Trim(), @"\(.*?\)", "").Trim();
                string valuePart = stripped[(ci+1)..].Trim();
                if (keyPart.StartsWith("Is he", StringComparison.OrdinalIgnoreCase) || keyPart.StartsWith("does he", StringComparison.OrdinalIgnoreCase)) continue;
                if (keyPart.Equals("Personal Information", StringComparison.OrdinalIgnoreCase)) continue;
                if (keyPart.StartsWith("L2", StringComparison.OrdinalIgnoreCase) && keyPart.Contains("Form", StringComparison.OrdinalIgnoreCase)) continue;
                string keyLow = keyPart.ToLower();
                // 짧은 키워드(예: "id")는 정확히 일치할 때만 매칭
                bool matched;
                if (kwLow.Length <= 2)
                    matched = keyLow == kwLow;
                else
                    matched = keyLow == kwLow || keyLow.Contains(kwLow) || kwLow.Contains(keyLow);
                if (matched)
                {
                    if (string.IsNullOrEmpty(valuePart) || valuePart.StartsWith("("))
                    { int lc = stripped.LastIndexOf(':'); if (lc > ci) valuePart = stripped[(lc+1)..].Trim(); }
                    if (!string.IsNullOrEmpty(valuePart)) return valuePart;
                }
            }
            if (kwLow == "id")
                foreach (string rawLine in lines)
                { string l = Regex.Replace(rawLine.Trim(), @"[()]", "").Trim();
                  if (Regex.IsMatch(l, @"^[A-Za-z]\d+[-\s]*\d+$")) return l.Replace(" ", ""); }
            return null;
        }

        private static string? SearchByNumberPattern(string[] lines, int number)
        {
            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim(); if (string.IsNullOrEmpty(line)) continue;
                var m = Regex.Match(line, @"^(\d+)\.\s*(.*)$"); if (!m.Success || int.Parse(m.Groups[1].Value) != number) continue;
                string rest = m.Groups[2].Value.Trim(); int ci = rest.IndexOf(':');
                if (ci >= 0) { string v = rest[(ci+1)..].Trim();
                    if (string.IsNullOrEmpty(v) || v.StartsWith("(")) { int lc = rest.LastIndexOf(':'); if (lc > ci) v = rest[(lc+1)..].Trim(); } return v; }
                else return rest;
            }
            return null;
        }

        private void BtnAddRow_Click(object? sender, EventArgs e) => AddTableRow();

        private void ResetScroll()
        {
            panelTable.AutoScrollPosition = new Point(0, 0);
            LayoutTable();
        }

        private void RemoveRows(List<int> indices)
        {
            foreach (int idx in indices.OrderByDescending(x => x))
            {
                panelTable.Controls.Remove(_txtKeywords[idx]);
                panelTable.Controls.Remove(_txtValues[idx]);
                panelTable.Controls.Remove(_rdoUnique[idx]);
                panelTable.Controls.Remove(_txtColNums[idx]);
                _txtKeywords[idx].Dispose();
                _txtValues[idx].Dispose();
                _rdoUnique[idx].Dispose();
                _txtColNums[idx].Dispose();

                var kwList = _txtKeywords.ToList(); kwList.RemoveAt(idx); _txtKeywords = kwList.ToArray();
                var valList = _txtValues.ToList(); valList.RemoveAt(idx); _txtValues = valList.ToArray();
                var rdoList = _rdoUnique.ToList(); rdoList.RemoveAt(idx); _rdoUnique = rdoList.ToArray();
                var colList = _txtColNums.ToList(); colList.RemoveAt(idx); _txtColNums = colList.ToArray();
            }
            panelTable.AutoScrollPosition = new Point(0, 0);
            LayoutTable();
            panelTable.Invalidate();
        }

        private void UniqueRadio_Click(int clickedIdx)
        {
            if (_rdoUnique[clickedIdx].Checked)
            {
                _rdoUnique[clickedIdx].Checked = false;
            }
            else
            {
                for (int i = 0; i < _rdoUnique.Length; i++)
                    _rdoUnique[i].Checked = false;
                _rdoUnique[clickedIdx].Checked = true;
            }
            SaveSettings();
        }

        private void AddTableRow()
        {
            int idx = _txtKeywords.Length;
            Array.Resize(ref _txtKeywords, idx + 1); Array.Resize(ref _txtValues, idx + 1); Array.Resize(ref _rdoUnique, idx + 1); Array.Resize(ref _txtColNums, idx + 1);
            Color CARD = Color.FromArgb(22,27,34); Color FG = Color.FromArgb(230,237,243); Color FG2 = Color.FromArgb(139,148,158);
            _txtKeywords[idx] = new TextBox { BackColor = CARD, ForeColor = FG2, Font = new Font("Consolas",8F), BorderStyle = BorderStyle.None, Text = "" };
            _txtKeywords[idx].Leave += (s, ev) => SaveSettings();
            _txtValues[idx] = new TextBox { BackColor = CARD, ForeColor = FG, Font = new Font("Segoe UI",9F), BorderStyle = BorderStyle.None, Text = "" };
            _rdoUnique[idx] = new RadioButton { BackColor = CARD, ForeColor = FG, AutoCheck = false, Appearance = Appearance.Normal, Text = "", Cursor = Cursors.Hand };
            int capturedIdx = idx;
            _rdoUnique[idx].Click += (s, ev) => UniqueRadio_Click(capturedIdx);
            _txtColNums[idx] = new TextBox { BackColor = Color.FromArgb(30,36,44), ForeColor = FG, Font = new Font("Consolas",9F), BorderStyle = BorderStyle.FixedSingle, TextAlign = HorizontalAlignment.Right, MaxLength = 3, Text = "" };
            _txtColNums[idx].KeyPress += (s, ev) => { if (!char.IsDigit(ev.KeyChar) && !char.IsControl(ev.KeyChar)) ev.Handled = true; };
            _txtColNums[idx].Leave += (s, ev) => { var tb = (TextBox)s!; if (tb.Text.Trim() == "0") tb.Text = ""; SaveSettings(); };
            panelTable.Controls.Add(_txtKeywords[idx]); panelTable.Controls.Add(_txtValues[idx]); panelTable.Controls.Add(_rdoUnique[idx]); panelTable.Controls.Add(_txtColNums[idx]);
            LayoutTable();
            panelTable.ScrollControlIntoView(_txtKeywords[idx]); _txtKeywords[idx].Focus();
        }

        private static string CleanFacebookAdd(string text)
        { text = Regex.Replace(text, @"\(\s*Facebook\s*Add[^)]*\)", "", RegexOptions.IgnoreCase);
          text = Regex.Replace(text, @"Facebook\s*Add\b[^,\r\n]*", "", RegexOptions.IgnoreCase); return text.Trim(); }

        private static string CleanSpecialChars(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new System.Text.StringBuilder(text.Length);
            foreach (char c in text)
            { if ((c>='A'&&c<='Z')||(c>='a'&&c<='z')||(c>='0'&&c<='9')) { sb.Append(c); continue; }
              if (c>=0xAC00&&c<=0xD7A3) { sb.Append(c); continue; } if (c>=0x3131&&c<=0x318E) { sb.Append(c); continue; }
              if (" -+/.@,:;'\"()_!?#&%$=~`<>[]{}\\|^".Contains(c)) { sb.Append(c); continue; } }
            return sb.ToString().Trim();
        }
    }
}
