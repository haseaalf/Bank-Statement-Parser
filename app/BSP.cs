// Bank Statement Parser (BSP) - standalone Windows application.
//
// Simple flow (like SLIK Parser): drop PDFs -> (optional) detect bank -> Proses -> tick results
// -> Unduh. No Excel/WPS needed: the window writes the job files, runs the R engine
// (BSP_modular_engine\BSP_engine.R) into a temporary folder and lets the user download the results.
// Built on the target PC by setup\build_exe.ps1 with the C# compiler that ships with
// Windows (.NET Framework 4.x, C# 5) - keep the code C# 5 compatible.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("Bank Statement Parser")]
[assembly: System.Reflection.AssemblyProduct("Bank Statement Parser (BSP)")]
[assembly: System.Reflection.AssemblyVersion("1.1.0.0")]

namespace BSP
{
    // ---------------------------------------------------------------- engine access
    static class Engine
    {
        public const string Entry = "BSP_engine.R";

        public static string AppDir { get { return AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\'); } }

        public static string WorkDir
        {
            get
            {
                string p = Path.Combine(Path.GetTempPath(), "BSP_work");
                Directory.CreateDirectory(p);
                return p;
            }
        }

        // Value from bsp_config.txt (written by Install_BSP.bat) next to the program
        public static string ConfigValue(string key)
        {
            string f = Path.Combine(AppDir, "bsp_config.txt");
            if (!File.Exists(f)) return "";
            foreach (string line in File.ReadAllLines(f))
            {
                if (line.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase))
                    return line.Substring(key.Length + 1).Trim();
            }
            return "";
        }

        public static string RscriptIn(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder)) return "";
            folder = folder.Trim().TrimEnd('\\');
            if (File.Exists(Path.Combine(folder, "Rscript.exe"))) return Path.Combine(folder, "Rscript.exe");
            if (File.Exists(Path.Combine(folder, "bin", "Rscript.exe"))) return Path.Combine(folder, "bin", "Rscript.exe");
            return "";
        }

        // bsp_config.txt, else the newest R under Program Files / per-user Programs / Documents
        public static string FindRFolder()
        {
            string cfg = ConfigValue("RSCRIPT_DIR");
            if (RscriptIn(cfg) != "") return cfg;
            string[] bases = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "R"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "R"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "R") };
            string best = ""; Version bestV = new Version(0, 0);
            foreach (string b in bases)
            {
                if (!Directory.Exists(b)) continue;
                foreach (string d in Directory.GetDirectories(b, "R-*"))
                {
                    Version v;
                    if (!Version.TryParse(Path.GetFileName(d).Substring(2), out v)) continue;
                    if (File.Exists(Path.Combine(d, "bin", "Rscript.exe")) && v > bestV) { bestV = v; best = Path.Combine(d, "bin"); }
                }
            }
            return best;
        }

        public static string FindEngineFolder()
        {
            string local = Path.Combine(AppDir, "BSP_modular_engine");
            if (File.Exists(Path.Combine(local, Entry))) return local;
            string cfg = ConfigValue("ENGINE_DIR");
            if (cfg != "" && File.Exists(Path.Combine(cfg, Entry))) return cfg;
            return local;
        }

        // Runs "<Rscript> BSP_engine.R <work> <engine> <mode>" and waits; false + message on failure
        public static bool Run(string rscript, string engine, string mode, out string error)
        {
            error = "";
            string done = Path.Combine(WorkDir, "engine_done.txt");
            if (File.Exists(done)) File.Delete(done);
            var psi = new ProcessStartInfo(rscript,
                string.Format("\"{0}\" \"{1}\" \"{2}\" {3}", Path.Combine(engine, Entry), WorkDir, engine, mode));
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            try
            {
                using (var p = Process.Start(psi)) { p.WaitForExit(); }
            }
            catch (Exception ex)
            {
                error = "R tidak bisa dijalankan: " + ex.Message;
                return false;
            }
            if (!File.Exists(done))
            {
                error = "Engine R tidak selesai. Lihat log: " + Path.Combine(WorkDir, "engine_log.txt");
                return false;
            }
            string[] lines = File.ReadAllLines(done);
            if (lines.Length > 0 && lines[0] == "OK") return true;
            error = "Engine R berhenti dengan error:\n" + (lines.Length > 1 ? lines[1] : "") +
                    "\nLog: " + Path.Combine(WorkDir, "engine_log.txt");
            return false;
        }

        public static string CsvField(string v)
        {
            return "\"" + (v ?? "").Replace("\"", "\"\"") + "\"";
        }

        public static void WriteText(string file, string text)
        {
            File.WriteAllText(file, text, new UTF8Encoding(false));
        }

        // CSV with header -> list of rows (dictionary by column name)
        public static List<Dictionary<string, string>> ReadCsv(string file)
        {
            var rows = new List<Dictionary<string, string>>();
            if (!File.Exists(file)) return rows;
            string text = File.ReadAllText(file, Encoding.UTF8);
            var records = new List<List<string>>();
            var cur = new List<string>();
            var sb = new StringBuilder();
            bool inQ = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (inQ)
                {
                    if (c == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i++; }
                        else inQ = false;
                    }
                    else sb.Append(c);
                }
                else if (c == '"') inQ = true;
                else if (c == ',') { cur.Add(sb.ToString()); sb.Clear(); }
                else if (c == '\r') { }
                else if (c == '\n') { cur.Add(sb.ToString()); sb.Clear(); records.Add(cur); cur = new List<string>(); }
                else sb.Append(c);
            }
            if (sb.Length > 0 || cur.Count > 0) { cur.Add(sb.ToString()); records.Add(cur); }
            if (records.Count == 0) return rows;
            List<string> head = records[0];
            for (int r = 1; r < records.Count; r++)
            {
                var d = new Dictionary<string, string>();
                for (int k = 0; k < head.Count; k++) d[head[k]] = k < records[r].Count ? records[r][k] : "";
                rows.Add(d);
            }
            return rows;
        }

        public static string Get(Dictionary<string, string> row, string key)
        {
            string v;
            return row.TryGetValue(key, out v) ? v : "";
        }
    }

    // ---------------------------------------------------------------- settings (per user)
    static class Settings
    {
        static string File_ { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BSP", "settings.ini"); } }

        public static Dictionary<string, string> Load()
        {
            var d = new Dictionary<string, string>();
            try
            {
                if (File.Exists(File_))
                    foreach (string l in File.ReadAllLines(File_, Encoding.UTF8))
                    {
                        int i = l.IndexOf('=');
                        if (i > 0) d[l.Substring(0, i)] = l.Substring(i + 1);
                    }
            }
            catch { }
            return d;
        }

        public static void Save(Dictionary<string, string> d)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(File_));
                File.WriteAllLines(File_, d.Select(kv => kv.Key + "=" + kv.Value), Encoding.UTF8);
            }
            catch { }
        }
    }

    // ---------------------------------------------------------------- look & feel (palette of the BSP icon)
    static class Theme
    {
        public static readonly Color Navy = ColorTranslator.FromHtml("#04225A");      // titles, text on light
        public static readonly Color Blue = ColorTranslator.FromHtml("#0146B3");      // main action, strip
        public static readonly Color BlueDark = ColorTranslator.FromHtml("#01368C");  // hover
        public static readonly Color Accent = ColorTranslator.FromHtml("#429BFD");    // highlight
        public static readonly Color Tint = ColorTranslator.FromHtml("#EAF2FE");      // hover / secondary buttons
        public static readonly Color TintHover = ColorTranslator.FromHtml("#D6E6FD");
        public static readonly Color Line = ColorTranslator.FromHtml("#B8C9DE");      // borders
        public static readonly Color Bg = ColorTranslator.FromHtml("#F3F6FA");        // window
        public static readonly Color Card = Color.White;
        public static readonly Color Ink = ColorTranslator.FromHtml("#1F2A3C");
        public static readonly Color Muted = ColorTranslator.FromHtml("#5B677A");
        public static readonly Color Ok = ColorTranslator.FromHtml("#1E7B34");
        public static readonly Color Warn = ColorTranslator.FromHtml("#9A5B00");
        public static readonly Color Err = ColorTranslator.FromHtml("#B3261E");
        public const string Font = "Segoe UI";

        public static GraphicsPath Rounded(Rectangle r, int radius)
        {
            var p = new GraphicsPath();
            int d = radius * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    // Flat colored button with hover and a clear disabled look
    class FlatButton : Button
    {
        readonly Color bg, fg, hover, disBg, disFg;

        public FlatButton(string text, string kind)
        {
            Text = text;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            AutoSize = true;
            Padding = new Padding(12, 4, 12, 4);
            Font = new Font(Theme.Font, 9.5F, FontStyle.Bold);
            if (kind == "primary") { bg = Theme.Blue; fg = Color.White; hover = Theme.BlueDark; disBg = ColorTranslator.FromHtml("#AFC4E6"); disFg = Color.White; }
            else if (kind == "secondary") { bg = Theme.Tint; fg = Theme.Blue; hover = Theme.TintHover; disBg = ColorTranslator.FromHtml("#F0F2F5"); disFg = ColorTranslator.FromHtml("#9AA3B0"); }
            else { bg = Theme.Bg; fg = Theme.Muted; hover = ColorTranslator.FromHtml("#E3E8EF"); disBg = Theme.Bg; disFg = ColorTranslator.FromHtml("#A5ADB8"); Font = new Font(Theme.Font, 9F); }
            FlatAppearance.MouseOverBackColor = hover;
            FlatAppearance.MouseDownBackColor = hover;
            Apply();
            EnabledChanged += (s, e) => Apply();
        }

        void Apply()
        {
            BackColor = Enabled ? bg : disBg;
            ForeColor = Enabled ? fg : disFg;
            Cursor = Enabled ? Cursors.Hand : Cursors.Default;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (Enabled) { base.OnPaint(e); return; }
            // WinForms paints disabled flat buttons grey on grey; draw our own disabled look
            e.Graphics.Clear(disBg);
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, disFg,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    // White card with a thin border
    class CardPanel : Panel
    {
        public CardPanel() { BackColor = Theme.Card; DoubleBuffered = true; Padding = new Padding(1); }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (var pen = new Pen(Theme.Line)) e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }
    }

    // Drop area: drag PDFs onto it, or click it to choose files
    class DropZone : Panel
    {
        public event Action<string[]> FilesChosen;
        public event Action BrowseRequested;
        bool active;
        public bool Busy;

        public DropZone()
        {
            DoubleBuffered = true;
            BackColor = Theme.Card;
            Cursor = Cursors.Hand;
            AllowDrop = true;
            MouseEnter += (s, e) => SetActive(true);
            MouseLeave += (s, e) => SetActive(false);
            Click += (s, e) => { if (!Busy && BrowseRequested != null) BrowseRequested(); };
            DragEnter += (s, e) =>
            {
                if (!Busy && e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effect = DragDropEffects.Copy; SetActive(true); }
            };
            DragLeave += (s, e) => SetActive(false);
            DragDrop += (s, e) =>
            {
                SetActive(false);
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && FilesChosen != null) FilesChosen(files);
            };
            Resize += (s, e) => Invalidate();
        }

        void SetActive(bool a) { active = a; Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            bool on = active && !Busy;
            g.Clear(on ? Theme.Tint : Theme.Card);
            var r = new Rectangle(6, 6, Width - 13, Height - 13);
            using (var pen = new Pen(on ? Theme.Blue : Theme.Line, 2) { DashPattern = new float[] { 4, 3 } })
            using (var path = Theme.Rounded(r, 8))
                g.DrawPath(pen, path);
            int cx = Width / 2, cy = Height / 2;
            var center = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            if (Busy)
            {
                using (var f1 = new Font(Theme.Font, 13F, FontStyle.Bold)) using (var b = new SolidBrush(Theme.Navy))
                    g.DrawString("Sedang memproses...", f1, b, cx, cy - 10, center);
                using (var f2 = new Font(Theme.Font, 9.5F)) using (var b = new SolidBrush(Theme.Muted))
                    g.DrawString("Mohon tunggu hingga semua file selesai.", f2, b, cx, cy + 16, center);
                return;
            }
            using (var f1 = new Font(Theme.Font, 13F, FontStyle.Bold)) using (var b = new SolidBrush(Theme.Navy))
                g.DrawString("Seret & lepas file PDF rekening koran di sini", f1, b, cx, cy - 26, center);
            using (var f2 = new Font(Theme.Font, 9.5F)) using (var b = new SolidBrush(Theme.Muted))
                g.DrawString("bisa banyak file, banyak bank, dan banyak rekening sekaligus", f2, b, cx, cy - 4, center);
            var btn = new Rectangle(cx - 75, cy + 11, 150, 28);
            using (var path = Theme.Rounded(btn, 6)) using (var b = new SolidBrush(on ? Theme.BlueDark : Theme.Blue))
                g.FillPath(b, path);
            using (var f3 = new Font(Theme.Font, 9.5F, FontStyle.Bold))
                g.DrawString("Pilih File PDF", f3, Brushes.White, cx, cy + 25, center);
        }
    }

    // ---------------------------------------------------------------- main window
    class MainForm : Form
    {
        const string DeveloperEmail = "haseaalfian6@gmail.com";
        const string ModePerFile = "Per file (1 PDF = 1 Excel)";
        const string ModeCons = "Gabung per rekening";

        // file list columns
        const int F_NO = 0, F_FILE = 1, F_BANK = 2, F_PWD = 3, F_DET = 4, F_SCORE = 5, F_ACC = 6, F_NAME = 7, F_PERIOD = 8, F_STATUS = 9, F_PATH = 10;
        // result list columns
        const int R_CHECK = 0, R_FILE = 1, R_FROM = 2, R_KIND = 3, R_STATUS = 4, R_DOWNLOAD = 5;

        DropZone drop;
        DataGridView gridFiles, gridResults;
        DataGridViewComboBoxColumn colBank;
        ComboBox cboMode;
        NumericUpDown numTol;
        CheckBox chkRecap;
        Label lblFileCount, lblStatus, lblSummary, lblEmpty, lblBanks;
        ProgressBar progress;
        RichTextBox log;
        FlatButton btnDetect, btnProcess, btnDownload, btnClear, btnRemove, btnBanks;
        readonly List<Button> busyButtons = new List<Button>();
        readonly string sessionDir = Path.Combine(Path.GetTempPath(), "BSP_session_" + Process.GetCurrentProcess().Id);
        int runNo;
        int countOk, countWarn, countErr;
        string lastInputFolder = "", lastDownloadFolder = "";
        public string LastMessage = "";
        public bool Silent;

        public MainForm()
        {
            Text = "Bank Statement Parser";
            Font = new Font(Theme.Font, 9F);
            BackColor = Theme.Bg;
            var wa = Screen.PrimaryScreen.WorkingArea;
            Size = new Size(Math.Min(1180, wa.Width - 40), Math.Min(940, wa.Height - 40));
            MinimumSize = new Size(Math.Min(900, Width), Math.Min(700, Height));
            StartPosition = FormStartPosition.CenterScreen;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(16, 12, 16, 0), BackColor = Theme.Bg };
            body.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));   // drop zone
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 42));     // file list
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // options + actions
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // status + download
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 58));     // results | log (side by side)

            drop = new DropZone { Dock = DockStyle.Fill, Margin = new Padding(0) };
            drop.BrowseRequested += SelectFiles;
            drop.FilesChosen += AddFiles;
            body.Controls.Add(drop, 0, 0);
            body.Controls.Add(BuildFileCard(), 0, 1);
            body.Controls.Add(BuildActionRow(), 0, 2);
            body.Controls.Add(BuildStatusRow(), 0, 3);
            var split = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 8, 0, 8), BackColor = Theme.Bg };
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            split.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            split.Controls.Add(BuildResultCard(), 0, 0);
            split.Controls.Add(BuildLogCard(), 1, 0);
            body.Controls.Add(split, 0, 4);

            Controls.Add(body);
            Controls.Add(BuildFooter());
            Controls.Add(new Panel { Dock = DockStyle.Top, Height = 4, BackColor = Theme.Blue });
            Controls.Add(BuildHeader());

            LoadSettings();
            FormClosing += (s, e) => SaveSettings();
            FormClosed += (s, e) => { try { if (Directory.Exists(sessionDir)) Directory.Delete(sessionDir, true); } catch { } };
            Shown += async (s, e) => { if (Rscript != "" && colBank.Items.Count <= 1) await RefreshBanksAsync(false); };
            UpdateButtons();
        }

        // ------------------------------------------------------------ layout pieces
        Control BuildHeader()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 78, BackColor = Theme.Card, Padding = new Padding(16, 8, 16, 8) };
            var logo = new PictureBox { Dock = DockStyle.Left, Width = 62, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Theme.Card };
            try
            {
                var s = typeof(MainForm).Assembly.GetManifestResourceStream("BSP.logo.png");
                if (s != null) logo.Image = Image.FromStream(s);
            }
            catch { }
            var texts = new Panel { Dock = DockStyle.Fill, Padding = new Padding(14, 6, 0, 0), BackColor = Theme.Card };
            var sub = new Label
            {
                Dock = DockStyle.Top, Height = 24, ForeColor = Theme.Muted, Font = new Font(Theme.Font, 9.5F),
                Text = "Ubah PDF rekening koran berbagai bank menjadi Excel standar, lengkap dengan verifikasi saldo dan rekap mutasi bulanan."
            };
            var title = new Label { Dock = DockStyle.Top, Height = 30, ForeColor = Theme.Navy, Font = new Font(Theme.Font, 15F, FontStyle.Bold), Text = "Konversi Rekening Koran ke Excel" };
            texts.Controls.Add(sub);
            texts.Controls.Add(title);
            header.Controls.Add(texts);
            header.Controls.Add(logo);
            return header;
        }

        Control CardHeader(string title, Label right, params Control[] buttons)
        {
            var head = new Panel { Dock = DockStyle.Top, Height = 34, BackColor = Theme.Card, Padding = new Padding(12, 4, 8, 3) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, WrapContents = false, BackColor = Theme.Card };
            foreach (var b in buttons) { b.Margin = new Padding(6, 0, 0, 0); flow.Controls.Add(b); }
            if (right != null) { right.AutoSize = true; right.ForeColor = Theme.Muted; right.Margin = new Padding(8, 6, 4, 0); flow.Controls.Add(right); flow.Controls.SetChildIndex(right, 0); }
            head.Controls.Add(new Label { Text = title, Dock = DockStyle.Left, AutoSize = true, ForeColor = Theme.Navy, Font = new Font(Theme.Font, 10.5F, FontStyle.Bold), Padding = new Padding(0, 4, 0, 0) });
            head.Controls.Add(flow);
            return head;
        }

        Control Separator() { return new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Theme.Line }; }

        Control BuildFileCard()
        {
            var card = new CardPanel { Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
            gridFiles = NewGrid();
            gridFiles.ReadOnly = false;
            gridFiles.EditMode = DataGridViewEditMode.EditOnEnter;
            gridFiles.DataError += (s, e) => { e.ThrowException = false; };
            gridFiles.Columns.Add(Col("No", 38, true));
            gridFiles.Columns.Add(Col("File PDF", 260, true));
            colBank = new DataGridViewComboBoxColumn { HeaderText = "Bank", Width = 115, FlatStyle = FlatStyle.Flat, DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox };
            colBank.Items.Add("AUTO");
            gridFiles.Columns.Add(colBank);
            gridFiles.Columns.Add(Col("Password", 85, false));
            gridFiles.Columns.Add(Col("Bank terdeteksi", 190, true));
            gridFiles.Columns.Add(Col("Skor", 45, true));
            gridFiles.Columns.Add(Col("No. Rekening", 125, true));
            gridFiles.Columns.Add(Col("Nama Rekening", 190, true));
            gridFiles.Columns.Add(Col("Periode", 160, true));
            gridFiles.Columns.Add(Col("Status", 150, true));
            gridFiles.Columns.Add(Col("Path", 10, true));
            gridFiles.Columns[F_PATH].Visible = false;
            gridFiles.Columns[F_BANK].ToolTipText = "AUTO = deteksi otomatis. Pilih bank bila 'Tidak dikenali' atau 'Ragu'.";
            gridFiles.Columns[F_PWD].ToolTipText = "Isi bila PDF dikunci password.";
            gridFiles.CellFormatting += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == F_PWD && e.Value != null && (string)e.Value != "")
                { e.Value = new string('*', ((string)e.Value).Length); e.FormattingApplied = true; }
            };
            gridFiles.SelectionChanged += (s, e) => UpdateButtons();

            lblFileCount = new Label { Text = "0 file" };
            btnRemove = new FlatButton("Hapus baris terpilih", "quiet");
            btnRemove.Click += (s, e) => RemoveSelected();
            btnClear = new FlatButton("Hapus daftar", "quiet");
            btnClear.Click += (s, e) => ClearFiles(true);
            busyButtons.Add(btnRemove); busyButtons.Add(btnClear);

            card.Controls.Add(gridFiles);
            card.Controls.Add(Separator());
            card.Controls.Add(CardHeader("Daftar file", lblFileCount, btnRemove, btnClear));
            return card;
        }

        Control BuildActionRow()
        {
            var row = new Panel { Dock = DockStyle.Fill, Height = 46, Margin = new Padding(0, 10, 0, 0), BackColor = Theme.Bg };
            var left = new FlowLayoutPanel { Dock = DockStyle.Left, AutoSize = true, WrapContents = false, BackColor = Theme.Bg, Padding = new Padding(0, 6, 0, 0) };
            cboMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200, Margin = new Padding(4, 2, 18, 0) };
            cboMode.Items.AddRange(new object[] { ModePerFile, ModeCons });
            cboMode.SelectedIndex = 0;
            numTol = new NumericUpDown { DecimalPlaces = 2, Increment = 0.01M, Minimum = 0, Maximum = 1000, Value = 0.01M, Width = 70, Margin = new Padding(4, 2, 18, 0) };
            chkRecap = new CheckBox { Text = "Buat rekap mutasi", Checked = true, AutoSize = true, ForeColor = Theme.Ink, Margin = new Padding(0, 4, 0, 0) };
            var tip = new ToolTip();
            tip.SetToolTip(cboMode, "Per file: 1 PDF menjadi 1 Excel. Gabung per rekening: semua bulan satu rekening menjadi 1 Excel.");
            tip.SetToolTip(numTol, "Selisih saldo yang masih dianggap cocok saat verifikasi.");
            tip.SetToolTip(chkRecap, "Tambahan 1 file rekap: sheet REKAP + 1 sheet per rekening, dengan cek saldo bulanan.");
            left.Controls.Add(Caption("Mode output"));
            left.Controls.Add(cboMode);
            left.Controls.Add(Caption("Toleransi saldo"));
            left.Controls.Add(numTol);
            left.Controls.Add(chkRecap);

            var right = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, WrapContents = false, BackColor = Theme.Bg, Padding = new Padding(0, 2, 0, 0) };
            btnDetect = new FlatButton("Deteksi Bank", "secondary") { Height = 36 };
            btnDetect.Click += async (s, e) => await DetectAsync();
            btnProcess = new FlatButton("Proses", "primary") { Height = 36, Margin = new Padding(8, 0, 0, 0), Padding = new Padding(26, 4, 26, 4) };
            btnProcess.Font = new Font(Theme.Font, 10.5F, FontStyle.Bold);
            btnProcess.Click += async (s, e) => await ProcessAsync();
            busyButtons.Add(btnDetect); busyButtons.Add(btnProcess);
            right.Controls.Add(btnDetect);
            right.Controls.Add(btnProcess);

            row.Controls.Add(left);
            row.Controls.Add(right);
            return row;
        }

        static Label Caption(string text)
        {
            return new Label { Text = text, AutoSize = true, ForeColor = Theme.Ink, Margin = new Padding(0, 6, 0, 0) };
        }

        Control BuildStatusRow()
        {
            var row = new Panel { Dock = DockStyle.Fill, Height = 44, Margin = new Padding(0, 8, 0, 0), BackColor = Theme.Bg };
            btnDownload = new FlatButton("Unduh yang Dicentang...", "primary") { Dock = DockStyle.Right, AutoSize = false, Width = 200 };
            btnDownload.Click += (s, e) => DownloadChecked(null);
            var left = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 2, 16, 0), BackColor = Theme.Bg };
            lblStatus = new Label { Dock = DockStyle.Top, Height = 20, ForeColor = Theme.Ink, Text = "Menunggu file. Seret PDF ke kotak di atas atau klik untuk memilih." };
            progress = new ProgressBar { Dock = DockStyle.Top, Height = 8, Style = ProgressBarStyle.Continuous, Value = 0 };
            left.Controls.Add(progress);
            left.Controls.Add(lblStatus);
            row.Controls.Add(left);
            row.Controls.Add(btnDownload);
            return row;
        }

        Control BuildResultCard()
        {
            var card = new CardPanel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 5, 0) };
            gridResults = NewGrid();
            gridResults.ReadOnly = false;
            gridResults.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "", Width = 32 });
            gridResults.Columns.Add(Col("File hasil", 260, true));
            gridResults.Columns.Add(Col("Dari", 200, true));
            gridResults.Columns.Add(Col("Jenis", 90, true));
            gridResults.Columns.Add(Col("Status", 105, true));
            gridResults.Columns.Add(new DataGridViewButtonColumn { HeaderText = "", Width = 72, FlatStyle = FlatStyle.Flat });
            gridResults.Columns[R_DOWNLOAD].DefaultCellStyle.BackColor = Theme.Tint;
            gridResults.Columns[R_DOWNLOAD].DefaultCellStyle.ForeColor = Theme.Blue;
            gridResults.Columns[R_DOWNLOAD].DefaultCellStyle.SelectionBackColor = Theme.TintHover;
            gridResults.Columns[R_DOWNLOAD].DefaultCellStyle.SelectionForeColor = Theme.Blue;
            gridResults.Columns[R_DOWNLOAD].DefaultCellStyle.Font = new Font(Theme.Font, 9F, FontStyle.Bold);
            gridResults.Columns[R_FILE].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            gridResults.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (e.ColumnIndex == R_DOWNLOAD) DownloadSingle(e.RowIndex);
                if (e.ColumnIndex == R_CHECK) gridResults.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            gridResults.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0 && e.ColumnIndex != R_DOWNLOAD) DownloadSingle(e.RowIndex); };
            lblEmpty = new Label
            {
                Text = "Belum ada file diproses. Hasil konversi akan muncul di sini.", AutoSize = false, Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter, ForeColor = Theme.Muted, BackColor = Theme.Card
            };
            lblSummary = new Label { Text = "" };
            card.Controls.Add(lblEmpty);
            card.Controls.Add(gridResults);
            card.Controls.Add(Separator());
            card.Controls.Add(CardHeader("Hasil", lblSummary));
            lblEmpty.BringToFront();
            return card;
        }

        Control BuildLogCard()
        {
            var card = new CardPanel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 0, 0) };
            log = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = ColorTranslator.FromHtml("#FAFBFD"),
                Font = new Font("Consolas", 8.5F), ForeColor = Theme.Ink, DetectUrls = false, WordWrap = true
            };
            card.Controls.Add(log);
            card.Controls.Add(Separator());
            card.Controls.Add(CardHeader("Log proses", new Label { Text = "Detail verifikasi tiap file" }));
            return card;
        }

        Control BuildFooter()
        {
            var footer = new Panel { Dock = DockStyle.Bottom, Height = 40, BackColor = Theme.Bg, Padding = new Padding(16, 6, 16, 8) };
            var right = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, WrapContents = false, BackColor = Theme.Bg };
            lblBanks = new Label { AutoSize = true, ForeColor = Theme.Muted, Margin = new Padding(0, 6, 4, 0) };
            btnBanks = new FlatButton("Perbarui daftar bank", "quiet");
            btnBanks.Click += async (s, e) => await RefreshBanksAsync(true);
            var contact = new FlatButton("Kontak Developer", "quiet") { Margin = new Padding(6, 0, 0, 0) };
            contact.Click += (s, e) => ShowContact();
            busyButtons.Add(btnBanks);
            right.Controls.Add(lblBanks);
            right.Controls.Add(btnBanks);
            right.Controls.Add(contact);
            footer.Controls.Add(new Label
            {
                Dock = DockStyle.Fill, ForeColor = Theme.Muted, TextAlign = ContentAlignment.MiddleLeft,
                Text = "File diproses di komputer ini. Hasil sementara dihapus saat aplikasi ditutup - unduh hasil yang diperlukan."
            });
            footer.Controls.Add(right);
            return footer;
        }

        static DataGridView NewGrid()
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Theme.Card,
                BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false, GridColor = ColorTranslator.FromHtml("#E6ECF3"),
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal, ColumnHeadersHeight = 26,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            };
            g.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F1F5FB");
            g.ColumnHeadersDefaultCellStyle.ForeColor = Theme.Navy;
            g.ColumnHeadersDefaultCellStyle.Font = new Font(Theme.Font, 9F, FontStyle.Bold);
            g.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            g.DefaultCellStyle.SelectionBackColor = Theme.Tint;
            g.DefaultCellStyle.SelectionForeColor = Theme.Ink;
            g.DefaultCellStyle.ForeColor = Theme.Ink;
            g.RowTemplate.Height = 26;
            return g;
        }

        static DataGridViewTextBoxColumn Col(string header, int width, bool readOnly)
        {
            return new DataGridViewTextBoxColumn { HeaderText = header, Width = width, ReadOnly = readOnly, SortMode = DataGridViewColumnSortMode.NotSortable };
        }

        // ------------------------------------------------------------ settings
        void LoadSettings()
        {
            var s = Settings.Load();
            Func<string, string, string> get = (k, d) => { string v; return s.TryGetValue(k, out v) && v != "" ? v : d; };
            cboMode.SelectedItem = get("mode", ModePerFile) == ModeCons ? ModeCons : ModePerFile;
            decimal tol;
            if (decimal.TryParse(get("tolerance", "0.01"), NumberStyles.Any, CultureInfo.InvariantCulture, out tol)) numTol.Value = tol;
            chkRecap.Checked = get("recap_v2", "1") == "1";
            lastInputFolder = get("last_input_folder", "");
            lastDownloadFolder = get("last_download_folder", "");
        }

        void SaveSettings()
        {
            if (Silent) return;
            var s = Settings.Load();
            s["mode"] = (string)cboMode.SelectedItem;
            s["tolerance"] = numTol.Value.ToString(CultureInfo.InvariantCulture);
            s["recap_v2"] = chkRecap.Checked ? "1" : "0";
            s["last_input_folder"] = lastInputFolder;
            s["last_download_folder"] = lastDownloadFolder;
            Settings.Save(s);
        }

        // R and engine are found automatically (bsp_config.txt from Install_BSP.bat, else standard places)
        string Rscript { get { return Engine.RscriptIn(Engine.FindRFolder()); } }
        string EngineDir { get { return Engine.FindEngineFolder(); } }

        // ------------------------------------------------------------ messages / state
        void Msg(string text, MessageBoxIcon icon)
        {
            LastMessage += text + "\n";
            if (!Silent) MessageBox.Show(this, text, "Bank Statement Parser", MessageBoxButtons.OK, icon);
        }

        void Status(string text) { lblStatus.Text = text; }

        void SetBusy(bool busy, string text)
        {
            foreach (var b in busyButtons) b.Enabled = !busy;
            drop.Busy = busy; drop.Invalidate();
            progress.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            if (!busy) progress.Value = 0;
            gridFiles.ReadOnly = busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
            Status(text);
            if (!busy) UpdateButtons();
        }

        void UpdateButtons()
        {
            bool hasFiles = gridFiles.Rows.Count > 0;
            btnDetect.Enabled = hasFiles; btnProcess.Enabled = hasFiles; btnClear.Enabled = hasFiles;
            btnRemove.Enabled = gridFiles.SelectedRows.Count > 0;
            btnDownload.Enabled = gridResults.Rows.Cast<DataGridViewRow>().Any(r => !string.IsNullOrEmpty(r.Tag as string));
            lblFileCount.Text = gridFiles.Rows.Count + " file";
            lblEmpty.Visible = gridResults.Rows.Count == 0;
        }

        void Log(string text, Color color)
        {
            log.SelectionStart = log.TextLength;
            log.SelectionLength = 0;
            log.SelectionColor = color;
            log.AppendText(text + "\n");
            log.SelectionColor = log.ForeColor;
            log.ScrollToCaret();
        }

        // ------------------------------------------------------------ file list
        void SelectFiles()
        {
            using (var d = new OpenFileDialog { Filter = "PDF rekening koran (*.pdf)|*.pdf", Multiselect = true, Title = "Pilih file PDF rekening koran" })
            {
                if (Directory.Exists(lastInputFolder)) d.InitialDirectory = lastInputFolder;
                if (d.ShowDialog(this) == DialogResult.OK) AddFiles(d.FileNames);
            }
        }

        public void AddFiles(string[] files)
        {
            var existing = new HashSet<string>(gridFiles.Rows.Cast<DataGridViewRow>().Select(r => ((string)r.Cells[F_PATH].Value ?? "").ToLowerInvariant()));
            int added = 0, skipped = 0;
            foreach (string f in files)
            {
                if (!f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) { skipped++; continue; }
                if (existing.Contains(f.ToLowerInvariant())) continue;
                int i = gridFiles.Rows.Add();
                var r = gridFiles.Rows[i];
                r.Cells[F_NO].Value = (i + 1).ToString();
                r.Cells[F_FILE].Value = Path.GetFileName(f);
                r.Cells[F_FILE].ToolTipText = f;
                r.Cells[F_BANK].Value = "AUTO";
                r.Cells[F_PWD].Value = "";
                r.Cells[F_STATUS].Value = "Belum diproses";
                r.Cells[F_STATUS].Style.ForeColor = Theme.Muted;
                r.Cells[F_PATH].Value = f;
                existing.Add(f.ToLowerInvariant());
                lastInputFolder = Path.GetDirectoryName(f);
                added++;
            }
            if (skipped > 0) Log(string.Format("[INFO] {0} file dilewati (bukan PDF).", skipped), Theme.Muted);
            if (added > 0) Status(string.Format("{0} file ditambahkan. Klik 'Deteksi Bank' untuk mengecek, atau langsung 'Proses'.", added));
            gridFiles.ClearSelection();
            UpdateButtons();
        }

        public void ClearFiles(bool ask)
        {
            if (gridFiles.Rows.Count == 0) return;
            if (ask && !Silent && MessageBox.Show(this, "Hapus semua file dari daftar?", "Bank Statement Parser",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            gridFiles.Rows.Clear();
            Status("Daftar file dikosongkan.");
            UpdateButtons();
        }

        void RemoveSelected()
        {
            foreach (DataGridViewRow r in gridFiles.SelectedRows.Cast<DataGridViewRow>().ToList()) gridFiles.Rows.Remove(r);
            for (int i = 0; i < gridFiles.Rows.Count; i++) gridFiles.Rows[i].Cells[F_NO].Value = (i + 1).ToString();
            UpdateButtons();
        }

        // ------------------------------------------------------------ engine jobs
        bool CheckReady()
        {
            if (gridFiles.Rows.Count == 0)
            {
                Msg("Belum ada file PDF. Seret file ke kotak di atas atau klik untuk memilih.", MessageBoxIcon.Information);
                return false;
            }
            if (Rscript == "" || !File.Exists(Path.Combine(EngineDir, Engine.Entry)))
            {
                Msg("R atau engine BSP belum terpasang di komputer ini.\nJalankan Install_BSP.bat di folder program, lalu buka aplikasi lagi.", MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        void WriteManifest()
        {
            gridFiles.EndEdit();
            var sb = new StringBuilder("file_path,bank_id,variant,password,output_dir\n");
            foreach (DataGridViewRow r in gridFiles.Rows)
            {
                string bank = (string)r.Cells[F_BANK].Value ?? "";
                sb.Append(Engine.CsvField((string)r.Cells[F_PATH].Value)).Append(',')
                  .Append(Engine.CsvField(bank == "" ? "AUTO" : bank)).Append(",\"\",")
                  .Append(Engine.CsvField((string)r.Cells[F_PWD].Value ?? "")).Append(",\"\"\n");
            }
            Engine.WriteText(Path.Combine(Engine.WorkDir, "job_manifest.csv"), sb.ToString());
        }

        Task<string> RunEngineAsync(string mode)
        {
            string rs = Rscript, eng = EngineDir;
            return Task.Run(() => { string err; return Engine.Run(rs, eng, mode, out err) ? "" : err; });
        }

        DataGridViewRow FileRow(string path)
        {
            string p = (path ?? "").ToLowerInvariant();
            return gridFiles.Rows.Cast<DataGridViewRow>().FirstOrDefault(r => ((string)r.Cells[F_PATH].Value ?? "").ToLowerInvariant() == p);
        }

        static void SetStatus(DataGridViewCell c, string text, Color color)
        {
            c.Value = text;
            c.Style.ForeColor = color;
            c.Style.SelectionForeColor = color;
            c.Style.Font = new Font(Theme.Font, 9F, FontStyle.Bold);
        }

        // ------------------------------------------------------------ actions
        public async Task RefreshBanksAsync(bool report)
        {
            if (Rscript == "" || !File.Exists(Path.Combine(EngineDir, Engine.Entry)))
            {
                if (report) Msg("R atau engine BSP belum terpasang. Jalankan Install_BSP.bat terlebih dahulu.", MessageBoxIcon.Warning);
                return;
            }
            SetBusy(true, "Membaca daftar bank dari engine...");
            string err = await RunEngineAsync("list_banks");
            SetBusy(false, err == "" ? "Daftar bank diperbarui." : "Gagal membaca daftar bank.");
            if (err != "") { Msg(err, MessageBoxIcon.Error); return; }
            var rows = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "bank_registry.csv"));
            colBank.Items.Clear(); colBank.Items.Add("AUTO");
            foreach (string id in rows.Select(r => Engine.Get(r, "bank_id")).Distinct()) colBank.Items.Add(id);
            lblBanks.Text = (colBank.Items.Count - 1) + " bank didukung";
            if (report)
            {
                var list = string.Join("\n", rows.Select(r => "- " + Engine.Get(r, "bank_name") + "  (" + Engine.Get(r, "variant") + ")"));
                Msg("Daftar bank diperbarui. Format yang didukung:\n\n" + list, MessageBoxIcon.Information);
            }
        }

        public async Task DetectAsync()
        {
            if (!CheckReady()) return;
            if (colBank.Items.Count <= 1) await RefreshBanksAsync(false);
            WriteManifest();
            SetBusy(true, "Mendeteksi bank dari " + gridFiles.Rows.Count + " file...");
            string err = await RunEngineAsync("detect");
            if (err != "") { SetBusy(false, "Deteksi gagal."); Msg(err, MessageBoxIcon.Error); return; }
            int unknown = 0;
            foreach (var d in Engine.ReadCsv(Path.Combine(Engine.WorkDir, "detect_result.csv")))
            {
                var r = FileRow(Engine.Get(d, "file_path"));
                if (r == null) continue;
                string variant = Engine.Get(d, "variant");
                r.Cells[F_DET].Value = Engine.Get(d, "bank_id") + (variant != "" ? " - " + variant : "");
                r.Cells[F_SCORE].Value = Engine.Get(d, "score");
                FillMeta(r, d);
                string st = Engine.Get(d, "status");
                if (st == "AUTO" || st == "MANUAL") SetStatus(r.Cells[F_STATUS], st == "AUTO" ? "Terdeteksi" : "Dipilih manual", Theme.Ok);
                else if (st == "AMBIGUOUS") { SetStatus(r.Cells[F_STATUS], "Ragu - pilih bank", Theme.Warn); unknown++; }
                else if (st == "PASSWORD") { SetStatus(r.Cells[F_STATUS], "Perlu password", Theme.Err); unknown++; }
                else { SetStatus(r.Cells[F_STATUS], st == "UNKNOWN" ? "Tidak dikenali - pilih bank" : "Gagal dibaca", Theme.Err); unknown++; }
                r.Cells[F_STATUS].ToolTipText = Engine.Get(d, "message");
            }
            SetBusy(false, unknown == 0 ? "Semua bank terdeteksi. Klik 'Proses'."
                                        : unknown + " file perlu dicek: pilih bank / isi password di daftar file, lalu klik 'Proses'.");
        }

        static void FillMeta(DataGridViewRow r, Dictionary<string, string> d)
        {
            r.Cells[F_ACC].Value = Engine.Get(d, "account_no");
            r.Cells[F_NAME].Value = Engine.Get(d, "account_name");
            string ps = Engine.Get(d, "period_start");
            r.Cells[F_PERIOD].Value = ps != "" ? ps + " - " + Engine.Get(d, "period_end") : "";
        }

        public async Task ProcessAsync()
        {
            if (!CheckReady()) return;
            if (colBank.Items.Count <= 1) await RefreshBanksAsync(false);
            bool cons = (string)cboMode.SelectedItem == ModeCons;
            string outDir = Path.Combine(sessionDir, "run" + (++runNo));
            Directory.CreateDirectory(outDir);
            WriteManifest();
            var opt = new StringBuilder("key,value\n");
            opt.Append("output_dir,").Append(Engine.CsvField(outDir)).Append('\n');
            opt.Append("output_mode,").Append(cons ? "consolidated" : "per_file").Append('\n');
            opt.Append("on_exists,suffix\n");
            opt.Append("tolerance,").Append(numTol.Value.ToString(CultureInfo.InvariantCulture)).Append('\n');
            opt.Append("verify_only,0\n");
            opt.Append("name_style,").Append(cons ? "standard" : "original").Append('\n');
            opt.Append("recap,").Append(chkRecap.Checked ? "1" : "0").Append('\n');
            Engine.WriteText(Path.Combine(Engine.WorkDir, "job_options.csv"), opt.ToString());
            SaveSettings();

            int n = gridFiles.Rows.Count;
            SetBusy(true, "Memproses " + n + " file... (banyak file bisa memakan beberapa menit)");
            Log(string.Format("=== {0}  Proses {1} file  |  {2}{3}", DateTime.Now.ToString("HH:mm:ss"), n,
                              cons ? "gabung per rekening" : "per file", chkRecap.Checked ? " + rekap mutasi" : ""), Theme.Navy);
            string err = await RunEngineAsync("parse");
            if (err != "") { SetBusy(false, "Proses gagal."); Log("[GAGAL] " + err, Theme.Err); Msg(err, MessageBoxIcon.Error); return; }

            var res = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "run_result.csv"));
            var details = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "verification_detail.csv"));
            var cross = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "consolidation_detail.csv"));
            int ok = 0, warn = 0, bad = 0;
            foreach (var r in res)
            {
                string status = Engine.Get(r, "status"), name = Engine.Get(r, "file_name"), output = Engine.Get(r, "output_file");
                string group = Engine.Get(r, "group_id");
                bool isRecap = group == "REKAP", isCons = name.StartsWith("[GABUNGAN");
                // the file list shows the outcome per PDF
                var fr = FileRow(Engine.Get(r, "file_path"));
                if (fr != null)
                {
                    FillMeta(fr, r);
                    if (fr.Cells[F_DET].Value == null || (string)fr.Cells[F_DET].Value == "")
                        fr.Cells[F_DET].Value = Engine.Get(r, "bank_id") + (Engine.Get(r, "variant") != "" ? " - " + Engine.Get(r, "variant") : "");
                    SetStatus(fr.Cells[F_STATUS], status == "OK" ? "Selesai - OK" : status == "WARN" ? "Selesai - perlu dicek" : "Gagal",
                              status == "OK" ? Theme.Ok : status == "WARN" ? Theme.Warn : Theme.Err);
                    fr.Cells[F_STATUS].ToolTipText = Engine.Get(r, "message");
                }
                // results: everything that produced a file, plus PDFs that failed
                // per-file verification rows of a consolidated run have no file of their own: log only
                if (output == "" && status != "ERROR") { LogResult(r, details, cross, false, false); continue; }
                if (status == "OK") ok++; else if (status == "WARN") warn++; else bad++;
                AddResultRow(r, output, isRecap ? "Rekap mutasi" : isCons ? "Gabungan" : "Per file", status);
                LogResult(r, details, cross, isCons || isRecap, true);
            }
            countOk += ok; countWarn += warn; countErr += bad;
            lblSummary.Text = string.Format("{0} OK  ·  {1} perlu dicek  ·  {2} gagal", countOk, countWarn, countErr);
            SetBusy(false, string.Format("Selesai: {0} OK, {1} perlu dicek, {2} gagal. Centang hasil lalu klik 'Unduh yang Dicentang...'.", ok, warn, bad));
            LastMessage += lblStatus.Text + "\n";
        }

        void AddResultRow(Dictionary<string, string> r, string output, string kind, string status)
        {
            string name = Engine.Get(r, "file_name");
            int i = gridResults.Rows.Add();
            var row = gridResults.Rows[i];
            bool hasFile = output != "" && File.Exists(output);
            row.Tag = hasFile ? output : null;
            row.Cells[R_CHECK].Value = hasFile;
            row.Cells[R_CHECK].ReadOnly = !hasFile;
            row.Cells[R_FILE].Value = hasFile ? Path.GetFileName(output) : "(tidak ada file)";
            string from = name;
            if (kind == "Gabungan" || kind == "Rekap mutasi")
            {
                var src = Engine.Get(r, "source_files").Split(new[] { " | " }, StringSplitOptions.RemoveEmptyEntries);
                from = kind == "Rekap mutasi" ? name.Replace("[REKAP] ", "") : src.Length + " file: " + string.Join(", ", src);
                row.Cells[R_FROM].ToolTipText = string.Join("\n", src);
            }
            row.Cells[R_FROM].Value = from;
            row.Cells[R_FROM].Style.ForeColor = Theme.Muted;
            row.Cells[R_KIND].Value = kind;
            SetStatus(row.Cells[R_STATUS], status == "OK" ? "✓ OK" : status == "WARN" ? "! Perlu dicek" : "✕ Gagal",
                      status == "OK" ? Theme.Ok : status == "WARN" ? Theme.Warn : Theme.Err);
            row.Cells[R_STATUS].ToolTipText = Engine.Get(r, "message");
            row.Cells[R_DOWNLOAD].Value = hasFile ? "Unduh" : "";
            UpdateButtons();
        }

        void LogResult(Dictionary<string, string> r, List<Dictionary<string, string>> details, List<Dictionary<string, string>> cross, bool withCrossFile, bool withDetails)
        {
            string status = Engine.Get(r, "status"), name = Engine.Get(r, "file_name");
            string tag = status == "OK" ? "[OK]" : status == "WARN" ? "[WARNING]" : "[GAGAL]";
            Color c = status == "OK" ? Theme.Ok : status == "WARN" ? Theme.Warn : Theme.Err;
            string info = Engine.Get(r, "message");
            if (Engine.Get(r, "rows") != "")
                info = string.Format("{0} transaksi, saldo awal {1} ({2}) -> akhir {3}{4}", Engine.Get(r, "rows"), Money(Engine.Get(r, "opening")),
                    Engine.Get(r, "opening_source") == "stated" ? "dari PDF" : "dihitung", Money(Engine.Get(r, "closing")),
                    info != "" ? " | " + info : "");
            Log(string.Format("{0} {1}  {2}", tag, name, info), c);
            foreach (var d in details.Where(d => withDetails && Engine.Get(d, "file_name") == name))
            {
                string sev = Engine.Get(d, "severity");
                string where = Engine.Get(d, "excel_row") != "" ? "baris " + Engine.Get(d, "excel_row") + " (" + Engine.Get(d, "date") + "): " : "";
                Log("      " + (sev == "ERROR" ? "ERROR " : sev == "WARN" ? "WARN  " : "INFO  ") + where + Engine.Get(d, "message"),
                    sev == "ERROR" ? Theme.Err : sev == "WARN" ? Theme.Warn : Theme.Muted);
            }
            string gid = Engine.Get(r, "group_id");
            foreach (var x in cross.Where(x => withCrossFile && Engine.Get(x, "group_id") == gid && gid != ""))
            {
                string sev = Engine.Get(x, "severity");
                Log("      " + (sev == "ERROR" ? "ERROR " : "WARN  ") + Engine.Get(x, "file_name") + ": " + Engine.Get(x, "message"),
                    sev == "ERROR" ? Theme.Err : Theme.Warn);
            }
        }

        static string Money(string v)
        {
            double d;
            return double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d) ? d.ToString("#,##0.00", CultureInfo.InvariantCulture) : v;
        }

        // ------------------------------------------------------------ downloads
        void DownloadSingle(int row)
        {
            if (row < 0 || row >= gridResults.Rows.Count) return;
            string src = gridResults.Rows[row].Tag as string;
            if (string.IsNullOrEmpty(src) || !File.Exists(src)) return;
            using (var d = new SaveFileDialog { Title = "Simpan sebagai", FileName = Path.GetFileName(src), DefaultExt = "xlsx", Filter = "Excel (*.xlsx)|*.xlsx" })
            {
                if (Directory.Exists(lastDownloadFolder)) d.InitialDirectory = lastDownloadFolder;
                if (d.ShowDialog(this) != DialogResult.OK) return;
                File.Copy(src, d.FileName, true);
                lastDownloadFolder = Path.GetDirectoryName(d.FileName);
                Log("[INFO] Disimpan ke: " + d.FileName, Theme.Muted);
                Status("Tersimpan: " + d.FileName);
            }
        }

        public void DownloadChecked(string folder)
        {
            gridResults.EndEdit();
            var rows = gridResults.Rows.Cast<DataGridViewRow>()
                       .Where(r => !string.IsNullOrEmpty(r.Tag as string) && r.Cells[R_CHECK].Value is bool && (bool)r.Cells[R_CHECK].Value).ToList();
            if (rows.Count == 0) { Status("Belum ada hasil yang dicentang untuk diunduh."); return; }
            if (folder == null)
            {
                using (var d = new FolderBrowserDialog { Description = "Pilih folder tujuan untuk hasil yang dicentang" })
                {
                    if (Directory.Exists(lastDownloadFolder)) d.SelectedPath = lastDownloadFolder;
                    if (d.ShowDialog(this) != DialogResult.OK) return;
                    folder = d.SelectedPath;
                }
            }
            foreach (var r in rows)
            {
                string src = (string)r.Tag;
                string target = Path.Combine(folder, Path.GetFileName(src));
                string b = Path.Combine(folder, Path.GetFileNameWithoutExtension(src)), ext = Path.GetExtension(src);
                for (int k = 1; File.Exists(target); k++) target = b + "(" + k + ")" + ext;
                File.Copy(src, target);
            }
            lastDownloadFolder = folder;
            Log(string.Format("[INFO] {0} file (dicentang) disimpan ke: {1}", rows.Count, folder), Theme.Muted);
            Status(string.Format("{0} file disimpan ke: {1}", rows.Count, folder));
            LastMessage += lblStatus.Text + "\n";
            if (!Silent) Process.Start("explorer.exe", "\"" + folder + "\"");
        }

        // ------------------------------------------------------------ contact popup (bottom-right, like SLIK Parser)
        void ShowContact()
        {
            var p = new Form
            {
                Text = "Kontak Developer", FormBorderStyle = FormBorderStyle.FixedToolWindow, StartPosition = FormStartPosition.Manual,
                Size = new Size(320, 170), BackColor = Theme.Card, TopMost = true, ShowInTaskbar = false
            };
            var wa = Screen.PrimaryScreen.WorkingArea;
            p.Location = new Point(wa.Right - p.Width - 20, wa.Bottom - p.Height - 20);
            p.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 4, BackColor = Theme.Blue });
            var lbl = new Label { Text = "Hubungi developer via email:", AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Bounds = new Rectangle(0, 18, 304, 22), ForeColor = Theme.Ink };
            var link = new LinkLabel { Text = DeveloperEmail, AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Bounds = new Rectangle(0, 42, 304, 22), LinkColor = Theme.Blue };
            link.LinkClicked += (s, e) => { try { Process.Start("mailto:" + DeveloperEmail); } catch { } };
            var close = new FlatButton("Tutup", "secondary") { AutoSize = false, Bounds = new Rectangle(112, 78, 80, 30) };
            close.Click += (s, e) => p.Close();
            p.Controls.AddRange(new Control[] { lbl, link, close });
            p.Show(this);
        }

        // ------------------------------------------------------------ automated checks
        // BSP.exe --selftest <pdf1|pdf2|...> <download folder> <per_file|consolidated> <recap 0|1> <report file>
        public static int SelfTest(string[] a)
        {
            var sb = new StringBuilder();
            try
            {
                var f = new MainForm { Silent = true };
                f.CreateControl();
                f.cboMode.SelectedItem = a[3] == "consolidated" ? ModeCons : ModePerFile;
                f.chkRecap.Checked = a[4] == "1";
                sb.AppendLine("R=" + f.Rscript + " | engine=" + f.EngineDir);
                f.AddFiles(a[1].Split('|'));
                RunSync(f.RefreshBanksAsync(false));
                RunSync(f.DetectAsync());
                foreach (DataGridViewRow r in f.gridFiles.Rows)
                    sb.AppendLine("DETECT|" + r.Cells[F_FILE].Value + "|" + r.Cells[F_DET].Value + "|" + r.Cells[F_SCORE].Value + "|" + r.Cells[F_STATUS].Value);
                RunSync(f.ProcessAsync());
                foreach (DataGridViewRow r in f.gridResults.Rows)
                    sb.AppendLine("RESULT|" + r.Cells[R_FILE].Value + "|" + r.Cells[R_KIND].Value + "|" + r.Cells[R_STATUS].Value + "|" + r.Cells[R_CHECK].Value);
                foreach (DataGridViewRow r in f.gridFiles.Rows)
                    sb.AppendLine("FILE|" + r.Cells[F_FILE].Value + "|" + r.Cells[F_STATUS].Value);
                f.DownloadChecked(a[2]);
                sb.AppendLine("DOWNLOADED|" + string.Join("|", Directory.GetFiles(a[2]).Select(Path.GetFileName)));
                sb.AppendLine("LOG lines=" + f.log.Lines.Length);
                sb.Append(f.log.Text);
                sb.Append(f.LastMessage);
                File.WriteAllText(a[5], sb.ToString(), Encoding.UTF8);
                return 0;
            }
            catch (Exception ex)
            {
                sb.AppendLine("EXCEPTION " + ex);
                File.WriteAllText(a[5], sb.ToString(), Encoding.UTF8);
                return 1;
            }
        }

        // BSP.exe --screenshot <png prefix> <pdf1|pdf2|...> : renders the window empty, after detection and after processing
        public static int Screenshot(string[] a)
        {
            var f = new MainForm { Silent = true, StartPosition = FormStartPosition.Manual, Location = new Point(-3000, -3000), ShowInTaskbar = false };
            f.Show();
            Action<string> shot = name =>
            {
                Application.DoEvents();
                using (var bmp = new Bitmap(f.Width, f.Height))
                {
                    f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                    bmp.Save(a[1] + "_" + name + ".png", System.Drawing.Imaging.ImageFormat.Png);
                }
            };
            RunSync(f.RefreshBanksAsync(false));
            shot("1_kosong");
            f.AddFiles(a[2].Split('|'));
            RunSync(f.DetectAsync());
            shot("2_deteksi");
            RunSync(f.ProcessAsync());
            shot("3_hasil");
            f.Close();
            return 0;
        }

        // pump the message loop until the task finishes (continuations run on this UI thread)
        static void RunSync(Task t)
        {
            while (!t.IsCompleted) { Application.DoEvents(); System.Threading.Thread.Sleep(20); }
            if (t.IsFaulted) throw t.Exception.GetBaseException();
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length >= 6 && args[0] == "--selftest") return MainForm.SelfTest(args);
            if (args.Length >= 3 && args[0] == "--screenshot")
            {
                try { return MainForm.Screenshot(args); }
                catch (Exception ex) { File.WriteAllText(args[1] + "_error.txt", ex.ToString()); return 1; }
            }
            Application.Run(new MainForm());
            return 0;
        }
    }
}
