using System.Drawing.Drawing2D;

namespace DataParser
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        private TextBox txtInput = null!;
        private Button btnProcess = null!;
        private Button btnSelectExcel = null!;
        private Button btnInsert = null!;
        private Label lblInput = null!;
        private Label lblResult = null!;
        private Label lblExcelPath = null!;
        private Label lblStatus = null!;
        private Panel panelTop = null!;
        private Panel panelMid = null!;
        private Panel panelBottom = null!;
        private Panel panelTable = null!;

        // 결과 테이블 행 데이터
        private Label[] _lblTitles = null!;
        private Label[] _lblValues = null!;
        private TextBox[] _txtColNums = null!;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        // 라운드 버튼
        private class RoundButton : Button
        {
            public int Radius { get; set; } = 2;
            public RoundButton()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                e.Graphics.Clear(Parent?.BackColor ?? Color.Black);
                var rect = new Rectangle(1, 1, Width - 3, Height - 3);
                using var path = MakeRound(rect, Radius);
                using var bg = new SolidBrush(BackColor);
                e.Graphics.FillPath(bg, path);
                TextRenderer.DrawText(e.Graphics, Text, Font, new Rectangle(0, 0, Width, Height), ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            private static GraphicsPath MakeRound(Rectangle r, int rad)
            {
                var p = new GraphicsPath();
                int d = rad * 2;
                p.AddArc(r.X, r.Y, d, d, 180, 90);
                p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                p.CloseFigure();
                return p;
            }
        }

        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(int x1, int y1, int x2, int y2, int cx, int cy);

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            this.SuspendLayout();

            // === 색상 ===
            Color BG     = Color.FromArgb(13, 17, 23);
            Color CARD   = Color.FromArgb(22, 27, 34);
            Color BORDER = Color.FromArgb(48, 54, 61);
            Color FG     = Color.FromArgb(230, 237, 243);
            Color FG2    = Color.FromArgb(139, 148, 158);
            Color BLUE   = Color.FromArgb(88, 166, 255);
            Color GREEN  = Color.FromArgb(63, 185, 80);

            // ============================================================
            // 상단: 입력 영역
            // ============================================================
            panelTop = new Panel { Dock = DockStyle.Top, Height = 256, BackColor = BG, Padding = new Padding(10, 8, 10, 0) };

            lblInput = new Label
            {
                Text = "  ▌ DATA INPUT",
                ForeColor = BLUE,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 22
            };

            txtInput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = CARD,
                ForeColor = FG,
                BorderStyle = BorderStyle.FixedSingle,
                AcceptsReturn = true,
                AcceptsTab = true,
                WordWrap = true
            };

            var panelBtn1 = new Panel { Dock = DockStyle.Bottom, Height = 38, BackColor = BG };
            btnProcess = new RoundButton
            {
                Text = "▶  1차 가공",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = BLUE,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false
            };
            btnProcess.FlatAppearance.BorderSize = 0;
            btnProcess.Click += BtnProcess_Click;
            panelBtn1.Controls.Add(btnProcess);
            panelBtn1.Resize += (s, e) =>
            {
                var sz = TextRenderer.MeasureText(btnProcess.Text, btnProcess.Font);
                btnProcess.Size = new Size(sz.Width + 48, 32);
                btnProcess.Location = new Point((panelBtn1.Width - btnProcess.Width) / 2, (panelBtn1.Height - btnProcess.Height) / 2);
            };

            panelTop.Controls.Add(txtInput);
            panelTop.Controls.Add(panelBtn1);
            panelTop.Controls.Add(lblInput);

            // ============================================================
            // 중간: 결과 테이블 (Label + TextBox 기반)
            // ============================================================
            panelMid = new Panel { Dock = DockStyle.Fill, BackColor = BG, Padding = new Padding(10, 4, 10, 4) };

            lblResult = new Label
            {
                Text = "  ▌ PARSED RESULT",
                ForeColor = GREEN,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 22
            };

            int rowH = 24;
            int hdrH = 26;
            int count = 13;
            var hdrBg = Color.FromArgb(1, 4, 9);

            // 헤더 패널 (고정, 스크롤 안 됨)
            var panelHeader = new Panel { Dock = DockStyle.Top, Height = hdrH, BackColor = hdrBg };
            var hdrTitle = new Label { Text = "TITLE", BackColor = hdrBg, ForeColor = BLUE, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(6, 0, 0, 0) };
            var hdrValue = new Label { Text = "VALUE", BackColor = hdrBg, ForeColor = BLUE, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(6, 0, 0, 0) };
            var hdrCol = new Label { Text = "COL #", BackColor = hdrBg, ForeColor = BLUE, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            panelHeader.Controls.AddRange(new Control[] { hdrTitle, hdrValue, hdrCol });

            // 데이터 패널 (스크롤 가능)
            panelTable = new Panel { Dock = DockStyle.Fill, BackColor = CARD, AutoScroll = true };

            _lblTitles = new Label[count];
            _lblValues = new Label[count];
            _txtColNums = new TextBox[count];

            // 구분선 저장용
            var hLines = new Panel[count + 1]; // 가로선
            var vLine1 = new Panel { BackColor = BORDER, Width = 1 }; // title|value
            var vLine2 = new Panel { BackColor = BORDER, Width = 1 }; // value|col#

            for (int i = 0; i <= count; i++)
                hLines[i] = new Panel { BackColor = BORDER, Height = 1 };

            for (int i = 0; i < count; i++)
            {
                _lblTitles[i] = new Label
                {
                    BackColor = CARD, ForeColor = FG2,
                    Font = new Font("Consolas", 9F, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(6, 0, 0, 0)
                };
                _lblValues[i] = new Label
                {
                    BackColor = CARD, ForeColor = FG,
                    Font = new Font("Segoe UI", 9F),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(6, 0, 0, 0)
                };
                _txtColNums[i] = new TextBox
                {
                    BackColor = Color.FromArgb(30, 36, 44), ForeColor = FG,
                    Font = new Font("Consolas", 9F),
                    BorderStyle = BorderStyle.FixedSingle,
                    TextAlign = HorizontalAlignment.Right,
                    MaxLength = 3
                };
                _txtColNums[i].KeyPress += (s, ev) =>
                { if (!char.IsDigit(ev.KeyChar) && !char.IsControl(ev.KeyChar)) ev.Handled = true; };
                _txtColNums[i].Leave += (s, ev) =>
                { var tb = (TextBox)s!; if (tb.Text.Trim() == "0") tb.Text = ""; SaveSettings(); };

                panelTable.Controls.Add(_lblTitles[i]);
                panelTable.Controls.Add(_lblValues[i]);
                panelTable.Controls.Add(_txtColNums[i]);
            }

            // 구분선 컨트롤 추가 (맨 위에 올리기 위해 나중에 추가)
            for (int i = 0; i <= count; i++) panelTable.Controls.Add(hLines[i]);
            panelTable.Controls.Add(vLine1);
            panelTable.Controls.Add(vLine2);

            // 헤더 레이아웃
            panelHeader.Resize += (s, e) =>
            {
                int w = panelHeader.ClientSize.Width;
                int colW = 60;
                int titleW = (int)(w * 0.32);
                int valW = w - titleW - colW;
                hdrTitle.SetBounds(0, 0, titleW, hdrH);
                hdrValue.SetBounds(titleW, 0, valW, hdrH);
                hdrCol.SetBounds(titleW + valW, 0, colW, hdrH);
            };

            // 데이터 레이아웃
            panelTable.Resize += (s, e) =>
            {
                int w = panelTable.ClientSize.Width;
                int colW = 60;
                int titleW = (int)(w * 0.32);
                int valW = w - titleW - colW;
                int totalH = count * rowH + count + 1; // 행 + 구분선

                int y = 0;
                for (int i = 0; i < count; i++)
                {
                    hLines[i].SetBounds(0, y, w, 1);
                    y += 1;
                    _lblTitles[i].SetBounds(0, y, titleW, rowH);
                    _lblValues[i].SetBounds(titleW, y, valW, rowH);
                    _txtColNums[i].SetBounds(titleW + valW + 4, y + 2, colW - 8, rowH - 4);
                    y += rowH;
                }
                hLines[count].SetBounds(0, y, w, 1);

                // 세로선
                vLine1.SetBounds(titleW, 0, 1, y + 1);
                vLine2.SetBounds(titleW + valW, 0, 1, y + 1);

                // 세로선을 맨 앞으로
                vLine1.BringToFront();
                vLine2.BringToFront();
            };

            panelMid.Controls.Add(panelTable);   // Fill (스크롤)
            panelMid.Controls.Add(panelHeader);  // Top (고정)
            panelMid.Controls.Add(lblResult);
            // ============================================================
            // 하단: 엑셀 선택 + 구분선 + 입력처리
            // ============================================================
            panelBottom = new Panel { Dock = DockStyle.Bottom, Height = 120, BackColor = BG, Padding = new Padding(10, 6, 10, 8) };

            btnSelectExcel = new RoundButton
            {
                Text = "📂  엑셀파일 선택",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = BLUE,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false
            };
            btnSelectExcel.FlatAppearance.BorderSize = 0;
            btnSelectExcel.Click += BtnSelectExcel_Click;

            lblExcelPath = new Label
            {
                Text = "선택된 파일 없음",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7F),
                AutoSize = false,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var separator = new Panel { BackColor = Color.FromArgb(48, 54, 61), Height = 1 };

            btnInsert = new RoundButton
            {
                Text = "⬇  입력처리",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = GREEN,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false
            };
            btnInsert.FlatAppearance.BorderSize = 0;
            btnInsert.Click += BtnInsert_Click;

            lblStatus = new Label
            {
                Text = "",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8F),
                AutoSize = false,
                Height = 18,
                TextAlign = ContentAlignment.MiddleCenter
            };

            panelBottom.Controls.Add(btnSelectExcel);
            panelBottom.Controls.Add(lblExcelPath);
            panelBottom.Controls.Add(separator);
            panelBottom.Controls.Add(lblStatus);
            panelBottom.Controls.Add(btnInsert);

            panelBottom.Resize += (s, e) =>
            {
                int cx = panelBottom.ClientSize.Width;
                int pad = panelBottom.Padding.Left;

                var szExcel = TextRenderer.MeasureText(btnSelectExcel.Text, btnSelectExcel.Font);
                btnSelectExcel.Size = new Size(szExcel.Width + 48, 30);
                btnSelectExcel.Location = new Point(pad, 6);

                lblExcelPath.Location = new Point(pad + btnSelectExcel.Width + 8, 10);
                lblExcelPath.Width = cx - pad * 2 - btnSelectExcel.Width - 12;

                separator.Width = cx - pad * 2;
                separator.Location = new Point(pad, btnSelectExcel.Bottom + 20);

                var szIns = TextRenderer.MeasureText(btnInsert.Text, btnInsert.Font);
                btnInsert.Size = new Size(szIns.Width + 48, 30);
                btnInsert.Location = new Point((cx - btnInsert.Width) / 2, separator.Bottom + 20);
            };

            // ============================================================
            // Form
            // ============================================================
            this.Controls.Add(panelMid);
            this.Controls.Add(panelBottom);
            this.Controls.Add(panelTop);

            this.Text = "YUNSEUL-0";
            this.BackColor = BG;
            this.Size = new Size(490, 700);
            this.MinimumSize = new Size(420, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.AutoScaleMode = AutoScaleMode.Font;
            this.DoubleBuffered = true;

            this.Resize += (s, e) => this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));
            this.HandleCreated += (s, e) => this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));

            this.ResumeLayout(false);
        }
    }
}
