using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        private static readonly string[] FixedTitles = new[]
        {
            "ID", "Networker Name", "Full Name", "Gender", "Birthday",
            "Contact", "Occupation", "City of Residence", "Church Name",
            "Denomination", "Church Position", "Married/Single", "Facebook address"
        };

        private static readonly Dictionary<string, string[]> TitleAliases = new()
        {
            ["ID"] = new[] { "id" },
            ["Networker Name"] = new[] { "networker name", "networker", "name of the person who introduced" },
            ["Full Name"] = new[] { "full name", "name" },
            ["Gender"] = new[] { "gender" },
            ["Birthday"] = new[] { "birthday", "date of birth", "birth date", "birthdate", "dob" },
            ["Contact"] = new[] { "contact", "phone", "mobile", "cell" },
            ["Occupation"] = new[] { "occupation", "job", "work" },
            ["City of Residence"] = new[] { "city of residence", "city", "residence" },
            ["Church Name"] = new[] { "church name", "church" },
            ["Denomination"] = new[] { "denomination" },
            ["Church Position"] = new[] { "church position", "position" },
            ["Married/Single"] = new[] { "married/single", "married / single", "marital status", "married", "single" },
            ["Facebook address"] = new[] { "facebook address", "facebook", "fb" }
        };

        private static readonly Dictionary<int, string> NumberToTitle = new()
        {
            [1] = "Full Name", [2] = "Gender", [3] = "Birthday", [4] = "Contact",
            [5] = "Occupation", [6] = "City of Residence", [7] = "Church Name",
            [8] = "Denomination", [9] = "Church Position", [10] = "Married/Single",
            [11] = "Facebook address"
        };

        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            EnableFormDrag();
            AddCloseButton();
            InitTable();
            LoadSettings();
        }

        private void InitTable()
        {
            for (int i = 0; i < FixedTitles.Length; i++)
                _lblTitles[i].Text = FixedTitles[i];
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

        private void AddCloseButton()
        {
            var btnClose = new Label
            {
                Text = "✕", ForeColor = Color.FromArgb(139, 148, 158), BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold), Size = new Size(28, 22),
                TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand
            };
            btnClose.Click += (s, e) => Close();
            btnClose.MouseEnter += (s, e) => btnClose.ForeColor = Color.FromArgb(248, 81, 73);
            btnClose.MouseLeave += (s, e) => btnClose.ForeColor = Color.FromArgb(139, 148, 158);
            this.Controls.Add(btnClose);
            btnClose.BringToFront();
            this.Resize += (s, e) => btnClose.Location = new Point(this.ClientSize.Width - 32, 4);
            btnClose.Location = new Point(this.ClientSize.Width - 32, 4);
        }

        // ── 설정 저장/로드 ──
        private void SaveSettings()
        {
            try
            {
                var lines = new List<string> { "EXCEL=" + (_excelPath ?? "") };
                for (int i = 0; i < FixedTitles.Length; i++)
                    lines.Add(FixedTitles[i] + "=" + _txtColNums[i].Text.Trim());
                File.WriteAllLines(SettingsFile, lines);
            }
            catch { }
        }

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsFile)) return;
                var saved = new Dictionary<string, string>();
                string? savedExcel = null;
                foreach (string line in File.ReadAllLines(SettingsFile))
                {
                    int eq = line.IndexOf('=');
                    if (eq < 0) continue;
                    string key = line.Substring(0, eq), val = line.Substring(eq + 1);
                    if (key == "EXCEL") savedExcel = val; else saved[key] = val;
                }
                if (!string.IsNullOrEmpty(savedExcel) && File.Exists(savedExcel))
                {
                    _excelPath = savedExcel;
                    lblExcelPath.Text = Path.GetFileName(_excelPath);
                    lblExcelPath.ForeColor = Color.FromArgb(88, 166, 255);
                }
                for (int i = 0; i < FixedTitles.Length; i++)
                    if (saved.TryGetValue(FixedTitles[i], out string? v))
                        _txtColNums[i].Text = v;
            }
            catch { }
        }

        // ── 1차 가공 ──
        private void BtnProcess_Click(object? sender, EventArgs e)
        {
            string raw = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(raw))
            { MessageBox.Show("데이터를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            DisplayData(ParseData(raw));
        }

        // ── 엑셀 선택 ──
        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Title = "엑셀 파일 선택", Filter = "Excel Files|*.xlsx;*.xls", RestoreDirectory = true };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _excelPath = dlg.FileName;
                lblExcelPath.Text = Path.GetFileName(_excelPath);
                lblExcelPath.ForeColor = Color.FromArgb(88, 166, 255);
                SaveSettings();
            }
        }

        // ── 입력처리 ──
        private void BtnInsert_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            { MessageBox.Show("먼저 엑셀 파일을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var entries = new List<(int col, string value)>();
            for (int i = 0; i < FixedTitles.Length; i++)
            {
                string val = _lblValues[i].Text.Trim();
                string numStr = _txtColNums[i].Text.Trim();
                if (!string.IsNullOrEmpty(numStr) && int.TryParse(numStr, out int colIdx) && colIdx > 0)
                    entries.Add((colIdx, CleanSpecialChars(val)));
            }

            if (entries.Count == 0)
            { MessageBox.Show("컬럼 번호(COL #)가 지정된 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (!entries.Any(x => !string.IsNullOrEmpty(x.value)))
            { MessageBox.Show("입력할 데이터가 없습니다.\n먼저 데이터를 입력하고 1차 가공을 해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            try
            {
                var fi = new FileInfo(_excelPath);
                using var pkg = new ExcelPackage(fi);
                var ws = pkg.Workbook.Worksheets[0];
                int totalRows = ws.Dimension?.End.Row ?? 0;
                int totalCols = ws.Dimension?.End.Column ?? 0;
                int newRow = totalRows + 1;
                for (int r = 1; r <= totalRows; r++)
                {
                    bool allEmpty = true;
                    for (int c = 1; c <= totalCols; c++)
                    {
                        var v = ws.Cells[r, c].Value;
                        if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) { allEmpty = false; break; }
                    }
                    if (allEmpty) { newRow = r; break; }
                }
                foreach (var (col, value) in entries)
                    ws.Cells[newRow, col].Value = value;
                pkg.Save();
                MessageBox.Show($"엑셀 시트1의 {newRow}행에 데이터를 입력했습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetAfterInsert();
            }
            catch (Exception ex)
            { MessageBox.Show($"엑셀 저장 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ResetAfterInsert()
        {
            txtInput.Text = "";
            for (int i = 0; i < FixedTitles.Length; i++)
                _lblValues[i].Text = "";
        }

        private void DisplayData(Dictionary<string, string> data)
        {
            for (int i = 0; i < FixedTitles.Length; i++)
                _lblValues[i].Text = data.ContainsKey(FixedTitles[i]) ? data[FixedTitles[i]] : "";
        }

        // ── 파싱 ──
        private static string CleanFacebookAdd(string text)
        {
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\(\s*Facebook\s*Add[^)]*\)", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            text = System.Text.RegularExpressions.Regex.Replace(text, @"Facebook\s*Add\b[^,\r\n]*", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return text.Trim();
        }

        private static string CleanSpecialChars(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new System.Text.StringBuilder(text.Length);
            foreach (char c in text)
            {
                // ASCII 알파벳, 숫자
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')) { sb.Append(c); continue; }
                // 한글 완성형
                if (c >= 0xAC00 && c <= 0xD7A3) { sb.Append(c); continue; }
                // 한글 자모
                if (c >= 0x3131 && c <= 0x318E) { sb.Append(c); continue; }
                // 허용 문장부호
                if (" -+/.@,:;'\"()_!?#&%$=~`<>[]{}\\|^".Contains(c)) { sb.Append(c); continue; }
                // 그 외 모두 삭제 (☐●○★ 등 유니코드 특수기호)
            }
            return sb.ToString().Trim();
        }

        private Dictionary<string, string> ParseData(string rawText)
        {
            var result = FixedTitles.ToDictionary(t => t, t => "");
            rawText = CleanFacebookAdd(rawText);
            string[] lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                line = System.Text.RegularExpressions.Regex.Replace(line, @"[()]", "").Trim();
                if (System.Text.RegularExpressions.Regex.IsMatch(line, @"^[A-Za-z]\d+[-\s]*\d+$"))
                { result["ID"] = line.Replace(" ", ""); break; }
            }

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                var mNet = System.Text.RegularExpressions.Regex.Match(line, @"Networker(?:\s*Name)?\s*:\s*(.+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (mNet.Success) { result["Networker Name"] = mNet.Groups[1].Value.Trim(); break; }
            }

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                var mNum = System.Text.RegularExpressions.Regex.Match(line, @"^(\d+)\.\s*(.*)$");
                if (mNum.Success)
                {
                    int num = int.Parse(mNum.Groups[1].Value);
                    string rest = mNum.Groups[2].Value.Trim();
                    int colonIdx = rest.IndexOf(':');
                    string value;
                    if (colonIdx >= 0)
                    {
                        string keyPart = rest.Substring(0, colonIdx).Trim();
                        value = rest.Substring(colonIdx + 1).Trim();
                        keyPart = System.Text.RegularExpressions.Regex.Replace(keyPart, @"\(.*?\)", "").Trim();
                        string? matchedByKey = FindBestMatchingTitle(keyPart);
                        if (matchedByKey != null && string.IsNullOrEmpty(result[matchedByKey]))
                        {
                            if (string.IsNullOrEmpty(value) || value.StartsWith("("))
                            { int lc = rest.LastIndexOf(':'); if (lc > colonIdx) value = rest.Substring(lc + 1).Trim(); }
                            result[matchedByKey] = value;
                            continue;
                        }
                    }
                    else { value = rest; }
                    if (NumberToTitle.TryGetValue(num, out string? titleByNum) && string.IsNullOrEmpty(result[titleByNum]))
                        result[titleByNum] = value;
                    continue;
                }

                {
                    int colonIdx = line.IndexOf(':');
                    if (colonIdx < 0) continue;
                    string keyPart = line.Substring(0, colonIdx).Trim();
                    string valuePart = line.Substring(colonIdx + 1).Trim();
                    keyPart = System.Text.RegularExpressions.Regex.Replace(keyPart, @"\(.*?\)", "").Trim();
                    if (keyPart.StartsWith("Is he", StringComparison.OrdinalIgnoreCase) || keyPart.StartsWith("does he", StringComparison.OrdinalIgnoreCase)) continue;
                    if (keyPart.Equals("Personal Information", StringComparison.OrdinalIgnoreCase)) continue;
                    if (keyPart.StartsWith("L2", StringComparison.OrdinalIgnoreCase) && keyPart.Contains("Form", StringComparison.OrdinalIgnoreCase)) continue;
                    string? matched = FindBestMatchingTitle(keyPart);
                    if (matched != null && string.IsNullOrEmpty(result[matched])) result[matched] = valuePart;
                }
            }
            return result;
        }

        private string? FindBestMatchingTitle(string input)
        {
            string low = input.ToLower().Trim();
            foreach (var kvp in TitleAliases) foreach (var a in kvp.Value) if (low == a) return kvp.Key;
            string? best = null; int bestScore = 0;
            foreach (var kvp in TitleAliases) foreach (var a in kvp.Value)
                if (low.Contains(a) || a.Contains(low)) { int sc = Math.Min(low.Length, a.Length); if (sc > bestScore) { bestScore = sc; best = kvp.Key; } }
            if (best != null) return best;
            int bestDist = int.MaxValue;
            foreach (var kvp in TitleAliases) foreach (var a in kvp.Value)
                { int d = Lev(low, a); int th = Math.Max(3, (int)(Math.Max(low.Length, a.Length) * 0.4)); if (d < bestDist && d <= th) { bestDist = d; best = kvp.Key; } }
            return best;
        }

        private static int Lev(string s, string t)
        {
            int n = s.Length, m = t.Length; var d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;
            for (int i = 1; i <= n; i++) for (int j = 1; j <= m; j++)
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + ((s[i - 1] == t[j - 1]) ? 0 : 1));
            return d[n, m];
        }
    }
}
