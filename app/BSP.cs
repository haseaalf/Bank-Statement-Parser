// Bank Statement Parser (BSP) - standalone Windows application.
//
// Same workflow as BSP_v1.xlsm, without Excel/WPS: the window writes the job files,
// runs the R engine (BSP_modular_engine\BSP_engine.R) and shows the results.
// Built on the target PC by setup\build_exe.ps1 with the C# compiler that ships with
// Windows (.NET Framework 4.x, C# 5) - keep the code C# 5 compatible.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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

    // ---------------------------------------------------------------- main window
    class MainForm : Form
    {
        const string LocChosen = "Folder pilihan";
        const string LocPdf = "Sama dengan folder PDF";
        const string ModePerFile = "Per file (1 PDF = 1 Excel)";
        const string ModeCons = "Gabung per rekening";
        const string ExistSuffix = "Beri akhiran (2) (3) dst";
        const string ExistOverwrite = "Timpa file lama";
        const string NameOriginal = "Sama dengan nama PDF";
        const string NameStandard = "Standar (BANK_NoRek_Periode)";

        static readonly Color Blue = Color.FromArgb(31, 78, 121);
        static readonly Color Green = Color.FromArgb(198, 239, 206);
        static readonly Color Yellow = Color.FromArgb(255, 235, 156);
        static readonly Color Red = Color.FromArgb(255, 199, 206);

        TextBox txtR, txtEngine, txtOut;
        ComboBox cboLoc, cboMode, cboExist, cboName;
        NumericUpDown numTol;
        CheckBox chkOpen, chkRecap;
        Label lblInfo;
        ToolStripStatusLabel statusText;
        ProgressBar progress;
        DataGridView gridFiles, gridResult, gridVerify, gridBanks;
        DataGridViewComboBoxColumn colBank, colFormat;
        TabControl tabs;
        TabPage tabResult, tabVerify;
        List<Button> actionButtons = new List<Button>();
        string lastInputFolder = "";
        public string LastMessage = "";
        public bool Silent;

        public MainForm()
        {
            Text = "Bank Statement Parser (BSP)";
            Font = new Font("Segoe UI", 9F);
            Width = 1250; Height = 780;
            MinimumSize = new Size(980, 600);
            StartPosition = FormStartPosition.CenterScreen;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            tabs = new TabControl { Dock = DockStyle.Fill };
            var tabMain = new TabPage("Proses");
            tabResult = new TabPage("Hasil");
            tabVerify = new TabPage("Verifikasi");
            var tabBanks = new TabPage("Daftar Bank");
            tabs.TabPages.AddRange(new[] { tabMain, tabResult, tabVerify, tabBanks });

            var status = new StatusStrip();
            statusText = new ToolStripStatusLabel("Siap") { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            status.Items.Add(statusText);
            progress = new ProgressBar { Style = ProgressBarStyle.Marquee, Visible = false, Width = 200, Dock = DockStyle.Right };

            BuildMainTab(tabMain);
            gridResult = NewGrid(true);
            gridResult.CellDoubleClick += (s, e) => OpenResultFile(e.RowIndex);
            tabResult.Controls.Add(gridResult);
            tabResult.Controls.Add(ResultButtons());
            gridVerify = NewGrid(true);
            tabVerify.Controls.Add(gridVerify);
            gridBanks = NewGrid(true);
            tabBanks.Controls.Add(gridBanks);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 24 };
            bottom.Controls.Add(progress);
            Controls.Add(tabs);
            Controls.Add(bottom);
            Controls.Add(status);

            LoadSettings();
            FormClosing += (s, e) => SaveSettings();
            // TabControl creates the controls of a page only when it is first shown; filling a grid on a
            // page that was never shown and showing it later fails ("Error creating window handle").
            // Visit every page once while the window opens so all grids exist before data arrives.
            Load += (s, e) =>
            {
                for (int i = tabs.TabPages.Count - 1; i >= 0; i--) tabs.SelectedIndex = i;
            };
            Shown += async (s, e) => { if (Engine.RscriptIn(txtR.Text) != "" && gridBanks.Rows.Count == 0) await RefreshBanksAsync(false); };
        }

        // ------------------------------------------------------------ layout
        void BuildMainTab(TabPage page)
        {
            var settings = new TableLayoutPanel { Dock = DockStyle.Top, Height = 250, ColumnCount = 4, Padding = new Padding(8) };
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 360));

            txtR = new TextBox { Dock = DockStyle.Fill };
            txtEngine = new TextBox { Dock = DockStyle.Fill };
            txtOut = new TextBox { Dock = DockStyle.Fill };
            cboLoc = Combo(LocChosen, LocPdf);
            cboMode = Combo(ModePerFile, ModeCons);
            cboExist = Combo(ExistSuffix, ExistOverwrite);
            cboName = Combo(NameOriginal, NameStandard);
            numTol = new NumericUpDown { DecimalPlaces = 2, Increment = 0.01M, Minimum = 0, Maximum = 1000, Value = 0.01M, Width = 90 };
            chkOpen = new CheckBox { Text = "Buka folder output setelah selesai", Checked = true, AutoSize = true };
            chkRecap = new CheckBox { Text = "Buat rekap mutasi (semua file: sheet REKAP + 1 sheet per rekening, cek saldo bulanan)", AutoSize = true };
            lblInfo = new Label { AutoSize = true, ForeColor = Color.Gray };

            int r = 0;
            AddRow(settings, r++, "Folder R (berisi Rscript.exe)", txtR, BrowseButton(txtR, "Pilih folder R (berisi Rscript.exe)"), null);
            AddRow(settings, r++, "Folder engine R", txtEngine, BrowseButton(txtEngine, "Pilih folder BSP_modular_engine"), lblInfo);
            AddRow(settings, r++, "Folder tujuan output", txtOut, BrowseButton(txtOut, "Pilih folder tujuan output"), null);
            AddRow(settings, r++, "Lokasi penyimpanan", cboLoc, null, null);
            AddRow(settings, r++, "Mode output", cboMode, null, null);
            AddRow(settings, r++, "Jika file output sudah ada", cboExist, null, null);
            AddRow(settings, r++, "Nama file output", cboName, null,
                   new Label { Text = "Mode gabung selalu memakai nama standar", AutoSize = true, ForeColor = Color.Gray });
            AddRow(settings, r++, "Toleransi selisih saldo", numTol, null, chkOpen);
            settings.Controls.Add(chkRecap, 1, r);
            settings.SetColumnSpan(chkRecap, 3);
            settings.RowCount = r + 1;

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(8, 6, 8, 0) };
            buttons.Controls.Add(ActionButton("Pilih File PDF...", (s, e) => SelectFiles()));
            buttons.Controls.Add(ActionButton("Deteksi Bank", async (s, e) => await DetectAsync()));
            var go = ActionButton("PROSES", async (s, e) => await ProcessAsync(false));
            go.Font = new Font(Font, FontStyle.Bold); go.BackColor = Color.FromArgb(198, 239, 206);
            buttons.Controls.Add(go);
            buttons.Controls.Add(ActionButton("Verifikasi saja", async (s, e) => await ProcessAsync(true)));
            buttons.Controls.Add(ActionButton("Hapus daftar", (s, e) => ClearFiles(true)));
            buttons.Controls.Add(ActionButton("Hapus baris terpilih", (s, e) => RemoveSelected()));
            buttons.Controls.Add(ActionButton("Perbarui daftar bank", async (s, e) => await RefreshBanksAsync(true)));

            var hint = new Label
            {
                Dock = DockStyle.Top, Height = 22, Padding = new Padding(10, 4, 0, 0), ForeColor = Color.Gray,
                Text = "Tarik file PDF ke tabel, atau klik 'Pilih File PDF...'. Kolom Bank: AUTO = deteksi otomatis, atau pilih bank. Isi Password bila PDF terkunci."
            };

            gridFiles = NewGrid(false);
            gridFiles.AllowUserToAddRows = false;
            gridFiles.AllowDrop = true;
            gridFiles.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            gridFiles.DragDrop += (s, e) => AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
            gridFiles.DataError += (s, e) => { e.ThrowException = false; };
            gridFiles.Columns.Add(TextCol("No", 40, true));
            gridFiles.Columns.Add(TextCol("File PDF", 360, true));
            colBank = new DataGridViewComboBoxColumn { HeaderText = "Bank", Width = 120, FlatStyle = FlatStyle.Flat };
            colBank.Items.Add("AUTO");
            gridFiles.Columns.Add(colBank);
            colFormat = new DataGridViewComboBoxColumn { HeaderText = "Format (opsional)", Width = 170, FlatStyle = FlatStyle.Flat };
            colFormat.Items.Add("");
            gridFiles.Columns.Add(colFormat);
            gridFiles.Columns.Add(TextCol("Password", 90, false));
            foreach (string h in new[] { "Bank terdeteksi", "Skor", "No. Rekening", "Nama Rekening", "Periode", "Status", "Pesan" })
                gridFiles.Columns.Add(TextCol(h, h == "Skor" ? 50 : (h == "Pesan" ? 260 : 140), true));

            page.Controls.Add(gridFiles);
            page.Controls.Add(hint);
            page.Controls.Add(buttons);
            page.Controls.Add(settings);
        }

        static void AddRow(TableLayoutPanel t, int row, string label, Control input, Control button, Control extra)
        {
            t.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
            t.Controls.Add(input, 1, row);
            if (button != null) t.Controls.Add(button, 2, row);
            if (extra != null) { extra.Anchor = AnchorStyles.Left; t.Controls.Add(extra, 3, row); }
        }

        static ComboBox Combo(params string[] items)
        {
            var c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
            c.Items.AddRange(items);
            c.SelectedIndex = 0;
            return c;
        }

        Button BrowseButton(TextBox target, string title)
        {
            var b = new Button { Text = "Browse...", Width = 80 };
            b.Click += (s, e) =>
            {
                using (var d = new FolderBrowserDialog { Description = title, SelectedPath = Directory.Exists(target.Text) ? target.Text : "" })
                {
                    if (d.ShowDialog(this) == DialogResult.OK)
                    {
                        target.Text = d.SelectedPath;
                        if (target == txtOut) cboLoc.SelectedItem = LocChosen;
                    }
                }
            };
            return b;
        }

        Button ActionButton(string text, EventHandler click)
        {
            var b = new Button { Text = text, AutoSize = true, Height = 30, Padding = new Padding(6, 0, 6, 0) };
            b.Click += click;
            actionButtons.Add(b);
            return b;
        }

        Panel ResultButtons()
        {
            var p = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(6) };
            var open = new Button { Text = "Buka file terpilih", AutoSize = true };
            open.Click += (s, e) => { if (gridResult.CurrentRow != null) OpenResultFile(gridResult.CurrentRow.Index); };
            var folder = new Button { Text = "Buka folder output", AutoSize = true };
            folder.Click += (s, e) => OpenOutputFolder();
            var log = new Button { Text = "Buka log engine", AutoSize = true };
            log.Click += (s, e) => OpenPath(Path.Combine(Engine.WorkDir, "engine_log.txt"));
            p.Controls.AddRange(new Control[] { open, folder, log,
                new Label { Text = "Double-click baris untuk membuka file output.", AutoSize = true, ForeColor = Color.Gray, Padding = new Padding(0, 6, 0, 0) } });
            return p;
        }

        static DataGridView NewGrid(bool readOnly)
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill, ReadOnly = readOnly, AllowUserToAddRows = false, AllowUserToDeleteRows = !readOnly,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White, BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false
            };
            g.ColumnHeadersDefaultCellStyle.BackColor = Blue;
            g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            g.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            g.ColumnHeadersHeight = 28;
            g.ShowCellToolTips = false;
            return g;
        }

        static DataGridViewTextBoxColumn TextCol(string header, int width, bool readOnly)
        {
            return new DataGridViewTextBoxColumn { HeaderText = header, Width = width, ReadOnly = readOnly };
        }

        // ------------------------------------------------------------ settings
        void LoadSettings()
        {
            var s = Settings.Load();
            Func<string, string, string> get = (k, d) => { string v; return s.TryGetValue(k, out v) && v != "" ? v : d; };
            txtR.Text = get("r_folder", "");
            if (Engine.RscriptIn(txtR.Text) == "") txtR.Text = Engine.FindRFolder();
            txtEngine.Text = get("engine_folder", "");
            if (!File.Exists(Path.Combine(txtEngine.Text, Engine.Entry))) txtEngine.Text = Engine.FindEngineFolder();
            txtOut.Text = get("output_folder", "");
            Select(cboLoc, get("location", LocChosen));
            Select(cboMode, get("mode", ModePerFile));
            Select(cboExist, get("on_exists", ExistSuffix));
            Select(cboName, get("name_style", NameOriginal));
            decimal tol;
            if (decimal.TryParse(get("tolerance", "0.01"), NumberStyles.Any, CultureInfo.InvariantCulture, out tol)) numTol.Value = tol;
            chkOpen.Checked = get("open_folder", "1") == "1";
            chkRecap.Checked = get("recap", "0") == "1";
            lastInputFolder = get("last_input_folder", "");
        }

        void SaveSettings()
        {
            if (Silent) return;
            var s = new Dictionary<string, string>();
            s["r_folder"] = txtR.Text; s["engine_folder"] = txtEngine.Text; s["output_folder"] = txtOut.Text;
            s["location"] = (string)cboLoc.SelectedItem; s["mode"] = (string)cboMode.SelectedItem;
            s["on_exists"] = (string)cboExist.SelectedItem; s["name_style"] = (string)cboName.SelectedItem;
            s["tolerance"] = numTol.Value.ToString(CultureInfo.InvariantCulture);
            s["open_folder"] = chkOpen.Checked ? "1" : "0"; s["recap"] = chkRecap.Checked ? "1" : "0";
            s["last_input_folder"] = lastInputFolder;
            Settings.Save(s);
        }

        static void Select(ComboBox c, string value)
        {
            int i = c.Items.IndexOf(value);
            c.SelectedIndex = i >= 0 ? i : 0;
        }

        // ------------------------------------------------------------ messages / busy state
        void Msg(string text, MessageBoxIcon icon)
        {
            LastMessage += text + "\n";
            if (!Silent) MessageBox.Show(this, text, "BSP", MessageBoxButtons.OK, icon);
        }

        bool Ask(string text)
        {
            if (Silent) return true;
            return MessageBox.Show(this, text, "BSP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        void SetBusy(bool busy, string text)
        {
            foreach (var b in actionButtons) b.Enabled = !busy;
            progress.Visible = busy;
            statusText.Text = text;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        // ------------------------------------------------------------ file list
        void SelectFiles()
        {
            using (var d = new OpenFileDialog { Filter = "PDF rekening koran (*.pdf)|*.pdf", Multiselect = true, Title = "Pilih satu atau lebih file rekening koran" })
            {
                if (Directory.Exists(lastInputFolder)) d.InitialDirectory = lastInputFolder;
                if (d.ShowDialog(this) == DialogResult.OK) AddFiles(d.FileNames);
            }
        }

        public void AddFiles(IEnumerable<string> files)
        {
            var existing = new HashSet<string>(gridFiles.Rows.Cast<DataGridViewRow>().Select(r => ((string)r.Cells[1].Value ?? "").ToLowerInvariant()));
            foreach (string f in files)
            {
                if (!f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) || existing.Contains(f.ToLowerInvariant())) continue;
                int i = gridFiles.Rows.Add();
                gridFiles.Rows[i].Cells[0].Value = (i + 1).ToString();
                gridFiles.Rows[i].Cells[1].Value = f;
                gridFiles.Rows[i].Cells[2].Value = "AUTO";
                gridFiles.Rows[i].Cells[3].Value = "";
                existing.Add(f.ToLowerInvariant());
                lastInputFolder = Path.GetDirectoryName(f);
            }
        }

        public void ClearFiles(bool ask)
        {
            if (ask && gridFiles.Rows.Count > 0 && !Ask("Hapus semua file dari daftar?")) return;
            gridFiles.Rows.Clear();
        }

        void RemoveSelected()
        {
            foreach (DataGridViewRow r in gridFiles.SelectedRows) gridFiles.Rows.Remove(r);
            for (int i = 0; i < gridFiles.Rows.Count; i++) gridFiles.Rows[i].Cells[0].Value = (i + 1).ToString();
        }

        // ------------------------------------------------------------ job files
        string Rscript { get { return Engine.RscriptIn(txtR.Text); } }
        string EngineDir { get { return txtEngine.Text.Trim().TrimEnd('\\'); } }

        bool CheckSettings(bool needOutput)
        {
            if (Rscript == "")
            {
                Msg("Folder R belum benar: pilih folder yang berisi Rscript.exe (contoh: C:\\Program Files\\R\\R-4.2.2\\bin).\nJalankan Install_BSP.bat bila R belum terpasang.", MessageBoxIcon.Warning);
                return false;
            }
            if (!File.Exists(Path.Combine(EngineDir, Engine.Entry)))
            {
                Msg("Folder engine R belum benar: pilih folder yang berisi " + Engine.Entry + ".", MessageBoxIcon.Warning);
                return false;
            }
            if (needOutput && (string)cboLoc.SelectedItem != LocPdf && !Directory.Exists(txtOut.Text.Trim()))
            {
                Msg("Mohon pilih folder tujuan output.", MessageBoxIcon.Warning);
                return false;
            }
            if (gridFiles.Rows.Count == 0)
            {
                Msg("Belum ada file PDF. Klik 'Pilih File PDF...' atau tarik file ke tabel.", MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        void WriteManifest()
        {
            var sb = new StringBuilder("file_path,bank_id,variant,password,output_dir\n");
            bool pdfFolder = (string)cboLoc.SelectedItem == LocPdf;
            foreach (DataGridViewRow r in gridFiles.Rows)
            {
                string f = (string)r.Cells[1].Value ?? "";
                string bank = (string)r.Cells[2].Value ?? "";
                if (bank == "") bank = "AUTO";
                sb.Append(Engine.CsvField(f)).Append(',').Append(Engine.CsvField(bank)).Append(',')
                  .Append(Engine.CsvField((string)r.Cells[3].Value ?? "")).Append(',')
                  .Append(Engine.CsvField((string)r.Cells[4].Value ?? "")).Append(',')
                  .Append(Engine.CsvField(pdfFolder ? Path.GetDirectoryName(f) : "")).Append('\n');
            }
            Engine.WriteText(Path.Combine(Engine.WorkDir, "job_manifest.csv"), sb.ToString());
        }

        void WriteOptions(bool verifyOnly)
        {
            bool cons = (string)cboMode.SelectedItem == ModeCons;
            string outDir = txtOut.Text.Trim().TrimEnd('\\');
            if (outDir == "") outDir = Engine.WorkDir;
            var sb = new StringBuilder("key,value\n");
            sb.Append("output_dir,").Append(Engine.CsvField(outDir)).Append('\n');
            sb.Append("output_mode,").Append(cons ? "consolidated" : "per_file").Append('\n');
            sb.Append("on_exists,").Append((string)cboExist.SelectedItem == ExistOverwrite ? "overwrite" : "suffix").Append('\n');
            sb.Append("tolerance,").Append(numTol.Value.ToString(CultureInfo.InvariantCulture)).Append('\n');
            sb.Append("verify_only,").Append(verifyOnly ? "1" : "0").Append('\n');
            sb.Append("name_style,").Append(cons || (string)cboName.SelectedItem == NameStandard ? "standard" : "original").Append('\n');
            sb.Append("recap,").Append(chkRecap.Checked ? "1" : "0").Append('\n');
            Engine.WriteText(Path.Combine(Engine.WorkDir, "job_options.csv"), sb.ToString());
        }

        Task<string> RunEngineAsync(string mode)
        {
            string rs = Rscript, eng = EngineDir;
            return Task.Run(() => { string err; return Engine.Run(rs, eng, mode, out err) ? "" : err; });
        }

        // ------------------------------------------------------------ actions
        public async Task RefreshBanksAsync(bool report)
        {
            if (Rscript == "" || !File.Exists(Path.Combine(EngineDir, Engine.Entry)))
            {
                if (report) Msg("Mohon pilih folder R dan folder engine terlebih dahulu.", MessageBoxIcon.Warning);
                return;
            }
            SetBusy(true, "Membaca daftar bank dari engine...");
            string err = await RunEngineAsync("list_banks");
            SetBusy(false, err == "" ? "Siap" : "Gagal membaca daftar bank");
            if (err != "") { Msg(err, MessageBoxIcon.Error); return; }
            LoadBankRegistry();
            if (report) Msg("Daftar bank diperbarui: " + (colBank.Items.Count - 1) + " bank tersedia.", MessageBoxIcon.Information);
        }

        void LoadBankRegistry()
        {
            var rows = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "bank_registry.csv"));
            if (rows.Count == 0) return;
            colBank.Items.Clear(); colBank.Items.Add("AUTO");
            colFormat.Items.Clear(); colFormat.Items.Add("");
            foreach (string id in rows.Select(r => Engine.Get(r, "bank_id")).Distinct()) colBank.Items.Add(id);
            foreach (string v in rows.Select(r => Engine.Get(r, "variant")).Distinct()) colFormat.Items.Add(v);
            FillGrid(gridBanks, rows, new[] { "bank_id", "bank_name", "variant", "description" },
                     new[] { "Kode Bank", "Nama Bank", "Format", "Keterangan" }, null);
            lblInfo.Text = string.Format("Engine: {0} bank, {1} format", colBank.Items.Count - 1, rows.Count);
        }

        public async Task DetectAsync()
        {
            if (!CheckSettings(false)) return;
            if (colBank.Items.Count <= 1) await RefreshBanksAsync(false);
            WriteManifest();
            SetBusy(true, "Mendeteksi bank...");
            string err = await RunEngineAsync("detect");
            SetBusy(false, err == "" ? "Deteksi selesai" : "Deteksi gagal");
            if (err != "") { Msg(err, MessageBoxIcon.Error); return; }
            ApplyDetectResult();
        }

        void ApplyDetectResult()
        {
            var rows = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "detect_result.csv"));
            foreach (DataGridViewRow g in gridFiles.Rows)
            {
                string f = ((string)g.Cells[1].Value ?? "").ToLowerInvariant();
                var d = rows.FirstOrDefault(r => Engine.Get(r, "file_path").ToLowerInvariant() == f);
                if (d == null) continue;
                string variant = Engine.Get(d, "variant");
                g.Cells[5].Value = Engine.Get(d, "bank_id") + (variant != "" ? " - " + variant : "");
                g.Cells[6].Value = Engine.Get(d, "score");
                g.Cells[7].Value = Engine.Get(d, "account_no");
                g.Cells[8].Value = Engine.Get(d, "account_name");
                string ps = Engine.Get(d, "period_start");
                g.Cells[9].Value = ps != "" ? ps + " - " + Engine.Get(d, "period_end") : "";
                string st = Engine.Get(d, "status");
                g.Cells[10].Value = StatusLabel(st);
                g.Cells[10].Style.BackColor = StatusColor(st);
                g.Cells[11].Value = Engine.Get(d, "message");
            }
        }

        static string StatusLabel(string s)
        {
            switch (s)
            {
                case "AUTO": return "Terdeteksi otomatis";
                case "MANUAL": return "Dipilih manual";
                case "AMBIGUOUS": return "Ragu - mohon cek";
                case "UNKNOWN": return "Tidak dikenali";
                case "PASSWORD": return "Perlu password";
                default: return s;
            }
        }

        static Color StatusColor(string s)
        {
            switch (s)
            {
                case "OK": case "AUTO": case "MANUAL": case "INFO": return Green;
                case "WARN": case "AMBIGUOUS": return Yellow;
                default: return Red;
            }
        }

        public async Task ProcessAsync(bool verifyOnly)
        {
            if (!CheckSettings(!verifyOnly)) return;
            if (!Ask((verifyOnly ? "Verifikasi " : "Proses ") + gridFiles.Rows.Count + " file PDF sekarang?")) return;
            WriteManifest();
            WriteOptions(verifyOnly);
            SaveSettings();
            SetBusy(true, "Memproses " + gridFiles.Rows.Count + " file... (file banyak bisa memakan beberapa menit)");
            string err = await RunEngineAsync("parse");
            SetBusy(false, err == "" ? "Selesai" : "Proses gagal");
            if (err != "") { Msg(err, MessageBoxIcon.Error); return; }

            var res = LoadResults();
            LoadVerification();
            int ok = res.Count(r => Engine.Get(r, "status") == "OK");
            int warn = res.Count(r => Engine.Get(r, "status") == "WARN");
            int bad = res.Count - ok - warn;
            tabs.SelectedTab = tabResult;
            Msg(string.Format("Selesai!\n\nOK      : {0}\nWARN    : {1} (periksa tab Verifikasi)\nERROR   : {2}{3}",
                              ok, warn, bad, bad > 0 ? " (periksa tab Hasil & Verifikasi)" : ""),
                bad > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            if (!verifyOnly && chkOpen.Checked && !Silent) OpenOutputFolder();
        }

        List<Dictionary<string, string>> LoadResults()
        {
            var rows = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "run_result.csv"));
            foreach (var r in rows)
            {
                string src = Engine.Get(r, "opening_source");
                r["opening_source"] = src == "stated" ? "Dari PDF" : (src == "derived" ? "Dihitung" : src);
                r["output_name"] = Path.GetFileName(Engine.Get(r, "output_file"));
            }
            FillGrid(gridResult, rows,
                new[] { "file_name", "bank_id", "variant", "account_no", "account_name", "currency", "period_start", "period_end",
                        "rows", "opening", "opening_source", "total_debit", "total_credit", "closing", "n_error", "n_warn",
                        "status", "output_name", "message", "source_files" },
                new[] { "File PDF", "Bank", "Format", "No. Rekening", "Nama Rekening", "Mata Uang", "Periode Awal", "Periode Akhir",
                        "Jumlah Transaksi", "Saldo Awal", "Sumber Saldo Awal", "Total Debit", "Total Kredit", "Saldo Akhir", "Error",
                        "Warning", "Status", "File Output", "Pesan", "File Sumber (gabungan)" },
                "status");
            for (int i = 0; i < rows.Count; i++) gridResult.Rows[i].Tag = Engine.Get(rows[i], "output_file");
            foreach (string c in new[] { "Saldo Awal", "Total Debit", "Total Kredit", "Saldo Akhir" }) FormatNumber(gridResult, c);
            FitColumns(gridResult);
            return rows;
        }

        void LoadVerification()
        {
            var rows = Engine.ReadCsv(Path.Combine(Engine.WorkDir, "verification_detail.csv"));
            foreach (var c in Engine.ReadCsv(Path.Combine(Engine.WorkDir, "consolidation_detail.csv")))
            {
                c["check"] = "antar-file: " + Engine.Get(c, "check");
                rows.Add(c);
            }
            FillGrid(gridVerify, rows,
                new[] { "file_name", "output_file", "excel_row", "date", "description", "debit", "credit", "balance", "check", "severity", "message" },
                new[] { "File PDF", "File Output", "Baris Excel", "Tanggal", "Deskripsi", "Debit", "Kredit", "Saldo", "Pemeriksaan", "Tingkat", "Keterangan" },
                "severity");
            foreach (string c in new[] { "Debit", "Kredit", "Saldo" }) FormatNumber(gridVerify, c);
            FitColumns(gridVerify);
            tabVerify.Text = rows.Count == 0 ? "Verifikasi" : "Verifikasi (" + rows.Count + ")";
        }

        static void FillGrid(DataGridView g, List<Dictionary<string, string>> rows, string[] keys, string[] heads, string statusKey)
        {
            g.Rows.Clear(); g.Columns.Clear();
            for (int k = 0; k < keys.Length; k++) g.Columns.Add("c" + k, heads[k]);
            foreach (var r in rows)
            {
                int i = g.Rows.Add(keys.Select(k => (object)Engine.Get(r, k)).ToArray());
                if (statusKey != null)
                {
                    int col = Array.IndexOf(keys, statusKey);
                    g.Rows[i].Cells[col].Style.BackColor = StatusColor(Engine.Get(r, statusKey));
                }
            }
            FitColumns(g);
        }

        static void FitColumns(DataGridView g)
        {
            g.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            foreach (DataGridViewColumn c in g.Columns) if (c.Width > 420) c.Width = 420;
        }

        static void FormatNumber(DataGridView g, string header)
        {
            foreach (DataGridViewColumn c in g.Columns)
            {
                if (c.HeaderText != header) continue;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                foreach (DataGridViewRow r in g.Rows)
                {
                    double v;
                    if (double.TryParse((string)r.Cells[c.Index].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out v))
                        r.Cells[c.Index].Value = v.ToString("#,##0.00", CultureInfo.InvariantCulture);
                }
            }
        }

        void OpenResultFile(int row)
        {
            if (row < 0 || row >= gridResult.Rows.Count) return;
            string f = gridResult.Rows[row].Tag as string;
            if (!string.IsNullOrEmpty(f) && File.Exists(f)) OpenPath(f);
        }

        void OpenOutputFolder()
        {
            string dir = (string)cboLoc.SelectedItem == LocPdf ? lastInputFolder : txtOut.Text.Trim();
            if (Directory.Exists(dir)) Process.Start("explorer.exe", "\"" + dir + "\"");
        }

        static void OpenPath(string path)
        {
            try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show("Tidak bisa membuka " + path + "\n" + ex.Message, "BSP"); }
        }

        // ------------------------------------------------------------ automated self-test
        // BSP.exe --selftest <pdf1|pdf2|...> <output folder> <per_file|consolidated> <recap 0|1> <report file>
        public static int SelfTest(string[] a)
        {
            var sb = new StringBuilder();
            try
            {
                var f = new MainForm { Silent = true };
                f.CreateControl();
                f.txtOut.Text = a[2];
                f.cboLoc.SelectedItem = LocChosen;
                f.cboMode.SelectedItem = a[3] == "consolidated" ? ModeCons : ModePerFile;
                f.cboExist.SelectedItem = ExistOverwrite;
                f.chkRecap.Checked = a[4] == "1";
                f.chkOpen.Checked = false;
                sb.AppendLine("R=" + f.txtR.Text + " | engine=" + f.txtEngine.Text);
                f.AddFiles(a[1].Split('|'));
                RunSync(f.RefreshBanksAsync(true));
                RunSync(f.DetectAsync());
                foreach (DataGridViewRow r in f.gridFiles.Rows)
                    sb.AppendLine("DETECT|" + r.Cells[5].Value + "|" + r.Cells[6].Value + "|" + r.Cells[10].Value);
                RunSync(f.ProcessAsync(false));
                foreach (DataGridViewRow r in f.gridResult.Rows)
                    sb.AppendLine("RESULT|" + r.Cells[0].Value + "|" + r.Cells[1].Value + "|" + r.Cells[8].Value + "|" + r.Cells[16].Value + "|" + r.Cells[17].Value);
                sb.AppendLine("VERIFY rows=" + f.gridVerify.Rows.Count + " | banks=" + f.gridBanks.Rows.Count);
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

        // BSP.exe --screenshot <png prefix> <pdf1|pdf2|...> : renders every tab after detection (layout check)
        public static int Screenshot(string[] a)
        {
            var f = new MainForm { Silent = true };
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-3000, -3000);
            f.ShowInTaskbar = false;
            f.Show();
            f.AddFiles(a[2].Split('|'));
            RunSync(f.RefreshBanksAsync(false));
            RunSync(f.DetectAsync());
            if (a.Length > 3) { f.txtOut.Text = a[3]; f.chkRecap.Checked = true; f.chkOpen.Checked = false; RunSync(f.ProcessAsync(false)); }
            // the normal user flow: switch between all tabs after processing, then draw each one
            for (int pass = 0; pass < 2; pass++)
                for (int t = 0; t < f.tabs.TabPages.Count; t++) { f.tabs.SelectedIndex = t; Application.DoEvents(); }
            for (int t = 0; t < f.tabs.TabPages.Count; t++)
            {
                f.tabs.SelectedIndex = t;
                Application.DoEvents();
                using (var bmp = new Bitmap(f.Width, f.Height))
                {
                    f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                    bmp.Save(a[1] + "_" + t + ".png", System.Drawing.Imaging.ImageFormat.Png);
                }
            }
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
                catch (Exception ex)
                {
                    var w = ex as System.ComponentModel.Win32Exception;
                    File.WriteAllText(args[1] + "_error.txt", (w != null ? "NativeErrorCode=" + w.NativeErrorCode + Environment.NewLine : "") + ex.ToString());
                    return 1;
                }
            }
            Application.Run(new MainForm());
            return 0;
        }
    }
}
