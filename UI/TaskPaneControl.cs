using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookClassifierAddIn5.ML;
using OutlookClassifierAddIn5.Data;
using OutlookClassifierAddIn5.Services; // kept for constructor signature compatibility


namespace OutlookClassifierAddIn5.UI
{
    internal sealed class BufferedDataGridView : System.Windows.Forms.DataGridView
    {
        public BufferedDataGridView()
        {
            this.DoubleBuffered = true;              // reduce flicker
            this.RowHeadersVisible = false;
            this.EnableHeadersVisualStyles = false;  // allow theming
            this.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.MultiSelect = true;
            this.AllowUserToAddRows = false;
            this.AllowUserToResizeRows = false;
            this.ReadOnly = true;
            this.AutoGenerateColumns = false;
            this.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        }
    }

    public partial class TaskPaneControl : UserControl
    {
        private readonly Outlook.Application _app;
        private readonly ModelService _ml;
        private readonly FeedbackStore _store;
        private readonly QueueService _queue;
        private readonly string _modelPath; // ADD THIS LINE

        // UI
        private TabControl tabs;
        private TabPage tabBatch;
        private TabPage tabSingle;
        private DataGridView grid; // batch tab (optional; styled so theme code is happy)
        private Label lblStats;    // batch footer (optional)
        private Label lblSubject, lblFrom, lblConf;
        private FlowLayoutPanel top3Panel;
        private ComboBox cmbFolders;
        private Button btnApprove;

        // Selection-driven state
        private CancellationTokenSource _showCts;
        private string _currentEntryId;
        private string _currentPredicted;
        private double _currentConf;
        private List<(string name, float p)> _currentTop3 = new List<(string name, float p)>();

        // Folder lists (full vs display)
        private List<string> _mailFolderPathsFull = new List<string>();
        private bool _suppressFilter;

        private List<FolderItem> _folderList = new List<FolderItem>();

        //Batching tab
        private Button btnApproveAll;
        private System.Collections.Generic.List<Services.QueueService.Item> _batchItems =
            new System.Collections.Generic.List<Services.QueueService.Item>();
        private const int BATCH_PAGE_SIZE = 50;
        private Button btnRemoveSelected;
        private readonly System.Collections.Generic.HashSet<string> _excludedEntryIds =
            new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);

        // Context menu for right-click exclude
        private ContextMenuStrip gridMenu;

        // Auto-approve feature (optional)
        private bool _autoApproveEnabled;
        private double _autoApproveThreshold = 0.92;

        // NEW: group-oriented batch UI
        private SplitContainer batchSplit;
        private TreeView folderTree;
        private string _selectedGroupKey = null;

        // NEW: groups keyed by predicted target folder (full Outlook path)
        private Dictionary<string, List<Services.QueueService.Item>> _groups =
            new Dictionary<string, List<Services.QueueService.Item>>(StringComparer.OrdinalIgnoreCase);

        private bool _initialized = false;

        // Cache pretty "From" strings by EntryId so we only resolve once per row
        private readonly Dictionary<string, string> _prettyFromCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public void SetAutoApprove(bool on) { _autoApproveEnabled = on; }
        
        // optional if we later want to tweak the threshold from UI:
        // public void SetAutoApproveThreshold(double t) { _autoApproveThreshold = t; }
        public TaskPaneControl(Outlook.Application app, ModelService ml, FeedbackStore store, QueueService queue, string modelPath)
        {
            _app = app;
            _ml = ml;
            _store = store;
            _queue = queue;
            _modelPath = modelPath; 

            InitializeComponent();
            BuildUi();
            ApplyLightTheme();
        }

        public void Init()
        {
            if (_initialized) return;
            _initialized = true;

            // Combo behavior (stable dropdown)
            cmbFolders.DropDownStyle = ComboBoxStyle.DropDown;
            cmbFolders.IntegralHeight = false;
            cmbFolders.MaxDropDownItems = 20;
            cmbFolders.DropDownHeight = 300;

            // TaskPaneControl ctor (end)
            RefreshFolderList();
            RebindGridFromBatchItems();
            UpdateBatchCountsUi();

            // ⬇️ Build queues, then refresh Batch UI
            _ = BuildQueuesAndRefreshAsync();

            // clear single view
            ClearSingleView("Select an email");
        }

        private void BuildUi()
        {
            try { 
                // --- polished layout ---
                this.SuspendLayout();
                this.Dock = DockStyle.Fill;
                this.AutoScaleMode = AutoScaleMode.Font; // DPI-friendly

                // Tabs
                tabs = new TabControl { Dock = DockStyle.Fill, Padding = new System.Drawing.Point(12, 6) };
                tabBatch = new TabPage("Batch") { Padding = new Padding(8) };
                tabSingle = new TabPage("Single") { Padding = new Padding(8) };
                tabs.TabPages.Add(tabBatch);
                tabs.TabPages.Add(tabSingle);
                this.Controls.Add(tabs);

                // === Batch tab ===
                batchSplit = new SplitContainer
                {
                    Dock = DockStyle.Fill,
                    Orientation = Orientation.Vertical,
                };

                // keep ratio once we have a size
                this.Load += (s, e) => ConfigureSplitSafely();
                batchSplit.SizeChanged += (s, e) => ConfigureSplitSafely();

                folderTree = new TreeView
                {
                    Dock = DockStyle.Fill,
                    HideSelection = false,
                    ShowNodeToolTips = true
                };
                folderTree.AfterSelect += (s, e) =>
                {
                    _selectedGroupKey = e.Node?.Tag as string;
                    RebindGridForGroup(_selectedGroupKey);
                };

                // Use buffered grid
                grid = new BufferedDataGridView
                {
                    Dock = DockStyle.Fill
                };
                grid.CellDoubleClick += Grid_CellDoubleClick;
                grid.ShowCellToolTips = true;
                grid.CellToolTipTextNeeded += Grid_CellToolTipTextNeeded;

                // Define columns ONCE
                grid.Columns.Clear();
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "EntryId",
                    DataPropertyName = "EntryId",
                    Visible = false
                });
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Subject",
                    HeaderText = "Subject",
                    DataPropertyName = "Subject",
                    FillWeight = 42 // relative width (sum ~100 across fill columns)
                });
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "From",
                    HeaderText = "From",
                    DataPropertyName = "From",
                    FillWeight = 25
                });
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Predicted",
                    HeaderText = "Predicted",
                    DataPropertyName = "Predicted",
                    FillWeight = 25
                });
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Confidence",
                    HeaderText = "Conf",
                    DataPropertyName = "Confidence",
                    Width = 64,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "0.00" }
                });

                // Subtle row look
                grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = System.Drawing.Color.FromArgb(248, 248, 248)
                };
                grid.RowTemplate.Height = 24;

                // Put controls into the split panels
                batchSplit.Panel1.Controls.Add(folderTree);
                batchSplit.Panel2.Controls.Add(grid);

                // Footer: stats (left) + buttons (right)
                lblStats = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };

                btnApproveAll = new Button { Text = "Approve Group (0)", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2) };
                btnApproveAll.Click += async (s, e) => { await ApproveAllInSelectedGroupAsync(); };

                btnRemoveSelected = new Button { Text = "Exclude Selected", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2) };
                btnRemoveSelected.Click += (s, e) => { RemoveSelectedFromBatch(); };

                // Right-aligned button area
                var footerButtons = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft,
                    WrapContents = false,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink
                };
                footerButtons.Controls.Add(btnApproveAll);
                footerButtons.Controls.Add(btnRemoveSelected);

                // Footer container (2 columns: stats | buttons)
                var footer = new TableLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Padding = new Padding(0, 6, 0, 0),
                    Height = 40,
                    ColumnCount = 2,
                    RowCount = 1
                };
                footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
                footer.Controls.Add(lblStats, 0, 0);
                footer.Controls.Add(footerButtons, 1, 0);

                // Compose the Batch tab
                tabBatch.Controls.Clear();
                tabBatch.Controls.Add(batchSplit);
                tabBatch.Controls.Add(footer);

                // Context menu (right-click)
                gridMenu = new ContextMenuStrip();
                gridMenu.Items.Add("Exclude selected from batch", null, (s, e) => RemoveSelectedFromBatch());
                gridMenu.Items.Add("Restore all excluded", null, (s, e) => { _excludedEntryIds.Clear(); RefreshQueues(); });
                grid.ContextMenuStrip = gridMenu;

                // Keyboard shortcut: Delete = exclude
                grid.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) { RemoveSelectedFromBatch(); e.Handled = true; } };

                // === Single tab ===
                var p = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 6,
                    Padding = new Padding(8),
                    AutoSize = false
                };
                p.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // subject
                p.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // from
                p.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // conf
                p.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // top3
                p.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // combo
                p.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // approve

                lblSubject = new Label { AutoSize = true, MaximumSize = new Size(1200, 0) }; // wraps
                lblFrom = new Label { AutoSize = true };
                lblConf = new Label { AutoSize = true };

                top3Panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    WrapContents = true,
                    Margin = new Padding(0, 6, 0, 6)
                };

                cmbFolders = new ComboBox
                {
                    Dock = DockStyle.Top,
                    DropDownStyle = ComboBoxStyle.DropDown,
                    IntegralHeight = false,
                    MaxDropDownItems = 20
                };
                cmbFolders.TextChanged += CmbFolders_TextChanged;
                cmbFolders.DropDown += (s, e) => ClampComboDropDown();

                btnApprove = new Button
                {
                    Text = "Approve",
                    Dock = DockStyle.Top,
                    Height = 32,
                    Enabled = false,
                    Margin = new Padding(0, 6, 0, 0)
                };
                btnApprove.Click += async (sender, args) => { await ApproveSelectedAsync(); };

                // Layout Single tab
                p.Controls.Add(lblSubject);
                p.Controls.Add(lblFrom);
                p.Controls.Add(lblConf);
                p.Controls.Add(top3Panel);
                p.Controls.Add(cmbFolders);
                p.Controls.Add(btnApprove);
                tabSingle.Controls.Add(p);

                // Finalize
                this.ResumeLayout(performLayout: true);

                // Keep a 28–30% splitter ratio on open and resize
                this.HandleCreated += (s, e) => AdjustBatchSplitter();
                this.Resize += (s, e) => AdjustBatchSplitter();

                // Initial data binding calls
                RebindGridFromBatchItems();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "UI init error");
                throw; // rethrow so you also see it in VS Output
            }

        }

        private void UI(Action a)
        {
            if (this.IsDisposed) return;
            if (this.InvokeRequired) { try { this.BeginInvoke(a); } catch { } }
            else { try { a(); } catch { } }
        }

        private void AdjustBatchSplitter()
        {
            try
            {
                if (batchSplit == null || batchSplit.Width <= 0) return;
                // left pane ~28% of the split width, within min sizes
                int target = (int)(batchSplit.Width * 0.28);
                int min = Math.Max(200, batchSplit.Panel1MinSize);
                int max = Math.Max(min, batchSplit.Width - batchSplit.Panel2MinSize - 50);
                batchSplit.SplitterDistance = Math.Min(Math.Max(target, min), max);
            }
            catch { /* ignore sizing exceptions */ }
        }

        private void ConfigureSplitSafely()
        {
            if (batchSplit == null) return;

            int width = batchSplit.ClientSize.Width;
            int splitter = batchSplit.SplitterWidth;

            if (width <= 0) return;

            // your desired mins
            int wantMin1 = 200;
            int wantMin2 = 300;

            // If the container is narrower than desired, scale mins down proportionally
            int availableForPanels = width - splitter;
            if (availableForPanels <= 0) return;

            int desiredTotal = wantMin1 + wantMin2;
            int min1 = wantMin1, min2 = wantMin2;

            if (availableForPanels < desiredTotal)
            {
                // scale both mins so their sum fits
                double scale = availableForPanels / (double)Math.Max(1, desiredTotal);
                min1 = Math.Max(0, (int)(wantMin1 * scale));
                min2 = Math.Max(0, (int)(wantMin2 * scale));
            }

            // Apply mins AFTER we know they’re legal
            try
            {
                batchSplit.Panel1MinSize = min1;
                batchSplit.Panel2MinSize = min2;
            }
            catch
            {
                // If a race hits here, relax to zero mins
                batchSplit.Panel1MinSize = 0;
                batchSplit.Panel2MinSize = 0;
            }

            // Now set a target distance (e.g., 28% of width) and clamp to [min1, width - splitter - min2]
            int target = (int)(width * 0.28);
            int maxLeft = width - splitter - min2;
            int clamped = Math.Max(min1, Math.Min(target, maxLeft));

            if (clamped >= min1 && clamped <= maxLeft)
            {
                try { batchSplit.SplitterDistance = clamped; } catch { /* ignore layout races */ }
            }
        }
        // ========= THEME / COLORS =========

        private static void SetColorsRecursive(Control root, Color back, Color fore)
        {
            if (root == null) return;
            bool isButton = root is Button || root is CheckBox || root is RadioButton;
            if (!isButton) root.BackColor = back;
            root.ForeColor = fore;

            foreach (Control c in root.Controls)
                SetColorsRecursive(c, back, fore);
        }

        private static void StyleGrid(DataGridView gv, Color back, Color fore, bool dark)
        {
            if (gv == null) return;
            gv.EnableHeadersVisualStyles = false;

            gv.BackgroundColor = back;
            gv.GridColor = dark ? Color.DimGray : Color.Gainsboro;

            gv.DefaultCellStyle.BackColor = back;
            gv.DefaultCellStyle.ForeColor = fore;
            gv.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            gv.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;

            gv.AlternatingRowsDefaultCellStyle.BackColor = dark ? Color.FromArgb(40, 40, 40) : Color.WhiteSmoke;

            gv.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            gv.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
        }

        private void ApplyLightTheme()
        {
            var back = SystemColors.Window;
            var fore = SystemColors.WindowText;
            SetColorsRecursive(this, back, fore);
            StyleGrid(grid, back, fore, false);

            // Make top-3 suggestion buttons readable
            foreach (var b in top3Panel.Controls.OfType<Button>())
            {
                b.UseVisualStyleBackColor = false;
                b.BackColor = SystemColors.ControlLight;
                b.ForeColor = SystemColors.ControlText;
            }

            // Inputs
            cmbFolders.BackColor = back;
            cmbFolders.ForeColor = fore;
        }

        // ========= FOLDER ENUMERATION & FILTERING =========

        private void RefreshFolderList()
        {
            _mailFolderPathsFull.Clear();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in FolderMap.GetAllFolderPaths(_app))
            {
                if (IsValidFilingFolder(p))
                {
                    try
                    {
                        var ns = _app.Session;
                        var folder = FolderMap.ResolveFolderByPath(p, ns);
                        if (folder != null && seen.Add(p)) // Only add if folder actually exists and not already added
                            _mailFolderPathsFull.Add(p);
                    }
                    catch
                    {
                        // Skip folders that can't be resolved (deleted/moved folders)
                    }
                }
            }

            var items = BuildFolderItems("");
            _folderList = items;

            // UI write: items / dropdown
            UI(() =>
            {
                _suppressFilter = true;
                cmbFolders.Items.Clear();
                foreach (var it in items)
                    if (!cmbFolders.Items.Contains(it))
                        cmbFolders.Items.Add(it);
                _suppressFilter = false;
            });
        }

        private sealed class FolderItem
        {
            public string Full { get; set; }    // full Outlook path
            public string Display { get; set; } // stripped for UX
            public override string ToString() { return Display; }
        }

        private List<FolderItem> BuildFolderItems(string filter)
        {
            var list = new List<FolderItem>();
            string f = (filter ?? string.Empty).ToLowerInvariant();

            // Precompute all (Full, Display) pairs
            var pairs = _mailFolderPathsFull
                .Select(full => new { Full = full, Display = StripAccountRoot(full) })
                .ToList();

            // Count how many share the same Display (case-insensitive)
            var displayCounts = pairs
                .GroupBy(p => p.Display, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

            foreach (var p in pairs)
            {
                // Apply filter across both display and full
                if (f.Length > 0 &&
                    !p.Display.ToLowerInvariant().Contains(f) &&
                    !p.Full.ToLowerInvariant().Contains(f))
                {
                    continue;
                }

                string disp = p.Display;

                // If duplicate display names exist, add account root to disambiguate
                if (displayCounts.TryGetValue(p.Display, out var dupCount) && dupCount > 1)
                {
                    // Account root = first segment before '/'
                    var firstSlash = p.Full.IndexOf('/');
                    var account = firstSlash > 0 ? p.Full.Substring(0, firstSlash) : p.Full;
                    disp = $"{p.Display} ({account})";
                }

                list.Add(new FolderItem { Full = p.Full, Display = disp });
            }

            // Optional: sort by display
            list.Sort((a, b) => string.Compare(a.Display, b.Display, StringComparison.OrdinalIgnoreCase));
            return list;
        }

        private List<FolderItem> GetFolderItems()
        {
            return _folderList;
        }

        private void CmbFolders_TextChanged(object sender, EventArgs e)
        {
            if (_suppressFilter) return;
            FuzzyFilterFolders();
        }

        private void FuzzyFilterFolders()
        {
            if (_suppressFilter) return;

            var text = cmbFolders.Text ?? string.Empty;
            var selStart = cmbFolders.SelectionStart;
            var selLen = cmbFolders.SelectionLength;
            bool wasFocused = cmbFolders.Focused;

            var filtered = BuildFolderItems(text);

            _suppressFilter = true;
            cmbFolders.BeginUpdate();
            cmbFolders.Items.Clear();
            foreach (var it in filtered) cmbFolders.Items.Add(it);
            cmbFolders.EndUpdate();

            // restore user text + caret
            cmbFolders.Text = text;
            cmbFolders.SelectionStart = selStart;
            cmbFolders.SelectionLength = selLen;

            // Only open if user is interacting
            cmbFolders.DroppedDown = wasFocused && filtered.Count > 0;

            // Size/width safely
            cmbFolders.IntegralHeight = false;
            cmbFolders.DropDownHeight = Math.Min(300, Math.Max(100, cmbFolders.ItemHeight * (filtered.Count + 1)));
            cmbFolders.DropDownWidth = MeasureDropDownWidth(cmbFolders, 40);

            _suppressFilter = false;
        }

        private static int MeasureDropDownWidth(ComboBox cmb, int extra)
        {
            int max = cmb.Width;
            using (var g = cmb.CreateGraphics())
            {
                foreach (var obj in cmb.Items)
                {
                    var s = obj == null ? string.Empty : obj.ToString();
                    var size = TextRenderer.MeasureText(g, s, cmb.Font);
                    if (size.Width > max) max = size.Width;
                }
            }
            var screen = Screen.FromControl(cmb).WorkingArea;
            max = Math.Min(max + SystemInformation.VerticalScrollBarWidth + extra, screen.Width - 40);
            return max;
        }
        public void RefreshQueues()
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)RefreshQueues); return; }
            if (_queue == null) return;

            var high = _queue.High.ToList();
            var med = _queue.Medium.ToList();
            var low = _queue.Low.ToList(); // ⬅️ new

            _batchItems = high.Concat(med)
                              .Where(m => !_excludedEntryIds.Contains(m.EntryId))
                              .OrderByDescending(m => m.Confidence)
                              .ToList();

            // ⬇️ Fallback: if no High/Medium, surface the top Low items (by confidence)
            if (_batchItems.Count == 0 && low.Count > 0)
            {
                _batchItems = low
                    .Where(m => !_excludedEntryIds.Contains(m.EntryId))
                    .OrderByDescending(m => m.Confidence)
                    .Take(BATCH_PAGE_SIZE) // reuse your page size
                    .ToList();
            }

            _groups = _batchItems
                .Select(m => new { Item = m, Target = GetEffectiveTarget(m) })
                .Where(x => !string.IsNullOrEmpty(x.Target))
                .GroupBy(x => x.Target, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(x => x.Item).OrderByDescending(i => i.Confidence).ToList(),
                    StringComparer.OrdinalIgnoreCase);

            BindFolderTree();

            int highCount = _batchItems.Count(x => x.Confidence >= _autoApproveThreshold);
            int medCount = _batchItems.Count - highCount;
            lblStats.Text = $"High: {highCount}  Medium: {medCount}  Excluded: {_excludedEntryIds.Count}  Folders: {_groups.Count}";
        }

        // ========= SINGLE-VIEW (SELECTION DRIVEN) =========

        public void BeginShowSelectedAsync(string entryId)
        {
            if (_showCts != null) _showCts.Cancel();
            _showCts = new CancellationTokenSource();
            var token = _showCts.Token;
            _ = ShowSingleForEntryIdAsync(entryId, token);
        }

        public void SelectSingleTab()
        {
            if (tabs != null && tabs.TabPages.Count > 1)
                tabs.SelectedIndex = 1;
        }

        public async Task ShowSingleForEntryIdAsync(string entryId, CancellationToken ct)
        {
            try
            {
                await Task.Yield();
                if (ct.IsCancellationRequested) return;

                var ns = _app.Session;
                var mail = ns.GetItemFromID(entryId) as Outlook.MailItem;
                if (mail == null)
                {
                    ClearSingleView("Select an email");
                    return;
                }

                string body = mail.Body ?? string.Empty;
                if (body.Length > 2000) body = body.Substring(0, 2000);
                var smtp = GetSenderSmtpAddress(mail);
                var row = new EmailRow
                {
                    Subject = mail.Subject ?? string.Empty,
                    Body = body,
                    FromAddress = mail.SenderEmailAddress ?? string.Empty,        // raw, as training had it
                    SenderDomain = ExtractDomain(mail.SenderEmailAddress ?? ""),  // derive from same string
                    HasAttachments = (mail.Attachments != null && mail.Attachments.Count > 0),
                    Label = string.Empty
                };

                if (ct.IsCancellationRequested) return;

                var result = _ml.PredictTop3(row);

                // Filter to valid filing folders: Inbox subfolders OR user-created folders (exclude system folders)
                var filtered = result.top3
                    .Where(t => IsValidFilingFolder(t.name))
                    .ToList();

                // Canonicalize predictions to full paths and filter to valid filing folders
                var top3Canon = new List<(string name, float p)>();
                foreach (var t in (result.top3 ?? new List<(string name, float p)>()))
                {
                    var full = CanonicalizeToFull(t.name);
                    if (!string.IsNullOrEmpty(full) && IsValidFilingFolder(full))
                        top3Canon.Add((full, t.p));
                }


                // Choose best full path; fall back to canonicalized primary prediction
                string predicted = top3Canon.Count > 0
                    ? top3Canon[0].name
                    : CanonicalizeToFull(result.folder);

                double conf = top3Canon.Count > 0 ? top3Canon[0].p : result.conf;

                // Auto-move only if user enabled it, confidence is high, and target is valid
                if (_autoApproveEnabled
                    && conf >= _autoApproveThreshold
                    && !string.IsNullOrEmpty(predicted)
                    && IsValidFilingFolder(predicted))
                {
                    await MoveAndLogSelectedAsync(entryId, predicted, predicted, conf);
                    ClearSingleView("Select an email");
                    return; // skip rendering the Single view since it's already moved
                }

                if (ct.IsCancellationRequested) return;

                // Update state
                _currentEntryId = entryId;
                _currentPredicted = predicted;
                _currentConf = conf;
                _currentTop3 = top3Canon;

                // UI update
                Action ui = () =>
                {
                    lblSubject.Text = row.Subject;

                    var senderName = GetSenderDisplayName(mail);
                    var pretty = !string.IsNullOrEmpty(smtp)
                        ? (string.IsNullOrEmpty(senderName) ? smtp : (senderName + " <" + smtp + ">"))
                        : (senderName ?? string.Empty);
                    lblFrom.Text = pretty;

                    lblConf.Text = "Conf: " + conf.ToString("0.00");

                    top3Panel.Controls.Clear();
                    foreach (var s in _currentTop3)
                    {
                        var display = ToShortDisplay(s.name); // short text for the button
                        var fullTarget = s.name;              // full path for the move
                        var b = new Button
                        {
                            Text = display + " (" + s.p.ToString("0.00") + ")",
                            AutoSize = true
                        };
                        var tip = new ToolTip();
                        tip.SetToolTip(b, StripAccountRoot(fullTarget));

                        b.Click += async (sender, args) =>
                        {
                            if (string.IsNullOrEmpty(fullTarget) || !IsValidFilingFolder(fullTarget))
                            {
                                MessageBox.Show("Choose a valid filing folder (not system folders like Inbox, Sent Items, etc.).", "Invalid folder");
                                return;
                            }

                            await MoveAndLogSelectedAsync(_currentEntryId, _currentPredicted, fullTarget, _currentConf);
                            ClearSingleView("Select an email");
                        };
                        top3Panel.Controls.Add(b);
                    }

                    // Refill dropdown without stealing focus; show short path text
                    _suppressFilter = true;
                    cmbFolders.Items.Clear();
                    foreach (var it in _folderList) cmbFolders.Items.Add(it);
                    cmbFolders.Text = StripAccountRoot(predicted);
                    _suppressFilter = false;

                    // Prefer focusing first suggestion, then Approve (not the ComboBox)
                    if (top3Panel.Controls.Count > 0) top3Panel.Controls[0].Focus();
                    else btnApprove.Focus();

                    btnApprove.Enabled = true;

                    // Ensure Single tab visible
                    if (tabs != null && tabs.TabPages.Count > 1)
                        tabs.SelectedIndex = 1;
                };

                if (this.InvokeRequired) this.BeginInvoke(ui); else ui();
            }
            catch
            {
                ClearSingleView("Select an email");
            }
        }

        private async Task ApproveSelectedAsync()
        {
            if (string.IsNullOrEmpty(_currentEntryId))
            {
                MessageBox.Show("No email selected.", "Approve");
                return;
            }

            string chosen = _currentPredicted; // fallback to model
            var sel = cmbFolders.SelectedItem as FolderItem;
            if (sel != null) chosen = sel.Full;
            else
            {
                var typed = cmbFolders.Text ?? string.Empty;
                var match = BuildFolderItems(typed).FirstOrDefault();
                if (match != null) chosen = match.Full;
                else
                    chosen = CanonicalizeToFull(typed); // map typed short/leaf to full path if unique
            }

            // NEW: ensure we move to a canonical full path
            chosen = CanonicalizeToFull(chosen);

            if (string.IsNullOrEmpty(chosen) || !IsValidFilingFolder(chosen))
            {
                MessageBox.Show("Choose a valid filing folder (not system folders like Inbox, Sent Items, etc.).", "Invalid folder");
                return;
            }

            await MoveAndLogSelectedAsync(_currentEntryId, _currentPredicted, chosen, _currentConf);
            ClearSingleView("Select an email");
        }

        private async System.Threading.Tasks.Task ApproveAllAsync()
        {
            if (_batchItems == null || _batchItems.Count == 0) return;

            int n = Math.Min(BATCH_PAGE_SIZE, _batchItems.Count);
            var page = _batchItems.Take(n).ToList();

            int moved = 0;
            foreach (var item in page)
            {
                // Use new validation logic
                string chosen = item.PredictedFolder;
                if (!IsValidFilingFolder(chosen))
                {
                    var alt = (item.Top3 ?? new System.Collections.Generic.List<(string name, float p)>())
                                .FirstOrDefault(t => IsValidFilingFolder(t.name));
                    if (string.IsNullOrEmpty(alt.name)) continue;
                    chosen = alt.name;
                }

                await MoveAndLogSelectedAsync(item.EntryId, item.PredictedFolder, chosen, item.Confidence);
                moved++;
            }

            if (_queue != null)
                await _queue.BuildQueuesAsync(_app);

            RefreshQueues();
            MessageBox.Show($"Approved {moved} item(s).", "Approve All");
        }


        private async System.Threading.Tasks.Task MoveAndLogSelectedAsync(
    string entryId, string predictedFolderPath, string chosenFolderPath, double confidence)
        {
            await SafeRun(async () =>
            {
                var ns = _app.Session;
                var mail = ns.GetItemFromID(entryId) as Outlook.MailItem;
                if (mail == null) return;

                // capture source path before move
                var sourceFolder = mail.Parent as Outlook.MAPIFolder;
                var sourcePath = GetFolderPath(sourceFolder);

                var target = FolderMap.ResolveFolderByPath(chosenFolderPath, ns);
                var moved = mail.Move(target) as Outlook.MailItem; // new EntryID after move
                var newId = moved != null ? moved.EntryID : entryId;

                // push undo record
                var rec = new MoveRecord
                {
                    OldEntryId = entryId,
                    NewEntryId = moved != null ? moved.EntryID : entryId,
                    SourcePath = sourcePath,
                    DestPath = chosenFolderPath,
                    Utc = DateTime.UtcNow
                };
                _undoStack.Push(rec);
                RefreshQueues();
                await _store.LogDecisionAsync(rec.NewEntryId, predictedFolderPath, chosenFolderPath, confidence);
                // remove from current batch page (if present) and update counts immediately
                this.BeginInvoke((Action)(() =>
                {
                    RemoveFromBatchById(entryId);  // old id on the page list
                    RemoveFromBatchById(newId);    // in case queues used the new id
                    RebindGridFromBatchItems();    // rebind -> row vanishes, counts update
                }));
            });
        }

        public void ClearSingleView(string message)
        {
            Action ui = () =>
            {
                _currentEntryId = null;
                _currentPredicted = null;
                _currentConf = 0;
                _currentTop3.Clear();

                lblSubject.Text = message ?? "Select an email";
                lblFrom.Text = string.Empty;
                lblConf.Text = string.Empty;
                top3Panel.Controls.Clear();

                // keep items for performance; just clear current text
                _suppressFilter = true;
                cmbFolders.Text = string.Empty;
                _suppressFilter = false;

                btnApprove.Enabled = false;
            };

            if (this.InvokeRequired) this.BeginInvoke(ui); else ui();
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var idObj = grid.Rows[e.RowIndex].Cells["EntryId"].Value;
            var entryId = idObj as string;
            if (string.IsNullOrEmpty(entryId)) return;

            OpenEmailByEntryId(entryId, openInspector: false);   // don't open window
                                                                // Also load it into the Single pane
            BeginShowSelectedAsync(entryId);
            SelectSingleTab();
        }

        private void OpenEmailByEntryId(string entryId, bool openInspector)
        {
            try
            {
                var ns = _app.Session;
                var mail = ns.GetItemFromID(entryId) as Outlook.MailItem;
                if (mail == null)
                {
                    MessageBox.Show("Couldn't locate the message.", "Open");
                    return;
                }

                if (openInspector)
                {
                    mail.Display(); // pop-out window
                    return;
                }

                // Explorer path (show in Reading Pane)
                var folder = mail.Parent as Outlook.MAPIFolder;

                // Use existing Explorer or create one on the correct folder
                var exp = _app.ActiveExplorer();
                if (exp == null)
                {
                    exp = _app.Explorers.Add(folder, Outlook.OlFolderDisplayMode.olFolderDisplayNormal);
                    exp.Display();
                }

                // Make sure we are in the right folder
                if (folder != null && exp.CurrentFolder != folder)
                    exp.CurrentFolder = folder;

                // Ensure Reading Pane is visible
                try
                {
                    if (!exp.IsPaneVisible(Outlook.OlPane.olPreview))
                        exp.ShowPane(Outlook.OlPane.olPreview, true);
                }
                catch { /* some configs throw here; safe to ignore */ }

                // Clear any existing selection, then select only this item
                try
                {
                    var currentSel = exp.Selection;
                    // Copy out before removing to avoid collection mutation issues
                    var toRemove = new System.Collections.Generic.List<object>();
                    for (int i = 1; i <= currentSel.Count; i++)
                        toRemove.Add(currentSel[i]);
                    foreach (var o in toRemove)
                        exp.RemoveFromSelection(o);
                }
                catch { /* best effort */ }

                exp.AddToSelection(mail);   // single-item selection updates Reading Pane
                exp.Activate();
            }
            catch (Exception ex)
            {
                // Fallback: if the Reading Pane path fails for any reason, open an inspector
                try { var ns = _app.Session; var mail = ns.GetItemFromID(entryId) as Outlook.MailItem; mail?.Display(); }
                catch { }
                MessageBox.Show(ex.Message, "Open");
            }
        }

        // ========= HELPERS =========

        public async Task SafeRun(Func<Task> fn)
        {
            try { await fn(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private static string ExtractDomain(string addr)
        {
            if (string.IsNullOrEmpty(addr)) return string.Empty;
            int at = addr.IndexOf('@');
            return at < 0 ? string.Empty : addr.Substring(at + 1).ToLowerInvariant();
        }

        private void Grid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            // For completeness; grid not actively used in selection mode
            e.ToolTipText = string.Empty;
        }

        public async System.Threading.Tasks.Task TrainModelAsync(string savePath)
        {
            await SafeRun(async () =>
            {
                // Use the new filtered method: recent data, no deleted items
                var rows = await _store.LoadTrainingRowsAsync(recentDays: 180, excludeDeletedItems: true);
                if (rows == null) return;

                _ml.Train(rows, out var _);
                if (!string.IsNullOrEmpty(savePath))
                    _ml.Save(savePath);

                // Re-score the currently selected email, if any
                if (!string.IsNullOrEmpty(_currentEntryId))
                    BeginShowSelectedAsync(_currentEntryId);
            });
        }

        private sealed class MoveRecord
        {
            public string OldEntryId;
            public string NewEntryId;
            public string SourcePath;
            public string DestPath;
            public DateTime Utc;
        }
        private readonly System.Collections.Generic.Stack<MoveRecord> _undoStack =
            new System.Collections.Generic.Stack<MoveRecord>();

        public void TryUndoLastMove()
        {
            if (_undoStack.Count == 0)
            {
                MessageBox.Show("Nothing to undo.", "Undo");
                return;
            }

            var rec = _undoStack.Pop();

            SafeRun(async () =>
            {
                var ns = _app.Session;

                // Try to find the item by the NEW id (post-move). Fall back to the old id.
                Outlook.MailItem item = null;
                try { item = ns.GetItemFromID(rec.NewEntryId) as Outlook.MailItem; } catch { }
                if (item == null)
                {
                    try { item = ns.GetItemFromID(rec.OldEntryId) as Outlook.MailItem; } catch { }
                }
                if (item == null)
                {
                    MessageBox.Show("Couldn’t find the message to undo.", "Undo");
                    return;
                }

                var src = FolderMap.ResolveFolderByPath(rec.SourcePath, ns);
                item.Move(src);

                // Optional: log the undo as a correction back to source
                await _store.LogDecisionAsync(rec.NewEntryId, rec.DestPath, rec.SourcePath, 1.0);

                ClearSingleView("Select an email");
            });
        }

        private void UpdateBatchCountsUi()
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)UpdateBatchCountsUi); return; }
            // how many will "Approve All" act on (this page only)
            int n = Math.Min(BATCH_PAGE_SIZE, _batchItems.Count);
            if (btnApproveAll != null)
            {
                btnApproveAll.Text = "Approve All (" + n + ")";
                btnApproveAll.Enabled = n > 0;
            }

            // derive High/Medium from current combined list (post-exclusions)
            int high = _batchItems.Count(x => x.Confidence >= 0.92);
            int med = _batchItems.Count - high;

            if (lblStats != null)
                lblStats.Text = "High: " + high + "  Medium: " + med + "  Excluded: " + _excludedEntryIds.Count + "  Total on page: " + n;
        }

        private void RebindGridFromBatchItems()
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)RebindGridFromBatchItems); return; }
            if (grid == null) return;

            var view = _batchItems.Select(m => new {
                m.EntryId,
                m.Subject,
                From = GetPrettyFrom(m.EntryId, m.From),
                Tier = (m.Confidence >= 0.92 ? "High" : "Medium"),
                PredictedFull = m.PredictedFolder,
                Predicted = ShortenPath(StripAccountRoot(m.PredictedFolder), 60),
                Confidence = m.Confidence.ToString("0.00")
            }).ToList();

            grid.DataSource = view;

            // make sure hidden key column exists
            if (grid.Columns["EntryId"] == null)
                grid.Columns.Insert(0, new DataGridViewTextBoxColumn { Name = "EntryId", DataPropertyName = "EntryId", Visible = false });
            else
                grid.Columns["EntryId"].Visible = false;

            UpdateBatchCountsUi();
        }

        private void RemoveFromBatchById(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return;
            _batchItems.RemoveAll(it => string.Equals(it.EntryId, entryId, StringComparison.Ordinal));
        }

        private static string GetFolderPath(Outlook.MAPIFolder f)
        {
            if (f == null) return string.Empty;
            string path = f.Name;
            var p = f.Parent as Outlook.MAPIFolder;
            while (p != null) { path = p.Name + "/" + path; p = p.Parent as Outlook.MAPIFolder; }
            return path;
        }

        private static string ShortenPath(string path, int maxLen = 48)
        {
            if (string.IsNullOrEmpty(path) || path.Length <= maxLen) return path;

            var parts = path.Split('/');
            var leaf = parts[parts.Length - 1];

            // Try to keep "…/Parent/Leaf" if it fits
            if (parts.Length >= 2)
            {
                var parent = parts[parts.Length - 2];
                var tail = parent + "/" + leaf;         // Parent/Leaf
                var withEllipsis = "…/" + tail;         // …/Parent/Leaf
                if (withEllipsis.Length <= maxLen)
                    return withEllipsis;
            }

            // Fall back to "…/Leaf"
            var leafOnly = "…/" + leaf;
            if (leafOnly.Length <= maxLen)
                return leafOnly;

            // If even that is too long, trim the leaf from the left (keep the end)
            int allowedLeaf = maxLen - 2; // account for "…/"
            if (allowedLeaf <= 0) return "…";
            return "…/" + leaf.Substring(leaf.Length - allowedLeaf);
        }

        private void RemoveSelectedFromBatch()
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)RemoveSelectedFromBatch); return; }
            if (grid.SelectedRows == null || grid.SelectedRows.Count == 0) return;

            var removedIds = new List<string>();
            foreach (DataGridViewRow r in grid.SelectedRows)
            {
                var id = r.Cells["EntryId"].Value as string;
                if (!string.IsNullOrEmpty(id) && _excludedEntryIds.Add(id))
                    removedIds.Add(id);
            }

            if (removedIds.Count == 0) return;

            // rebuild queues -> groups -> tree -> grid
            RefreshQueues();
        }

        private string GetEffectiveTarget(Services.QueueService.Item item)
        {
            // Always canonicalize the primary prediction to a full path if possible
            string target = CanonicalizeToFull(item.PredictedFolder);

            // If that target isn't a valid filing folder, try Top-3 alternatives (also canonicalized)
            if (!IsValidFilingFolder(target))
            {
                foreach (var t in (item.Top3 ?? new List<(string name, float p)>()))
                {
                    var alt = CanonicalizeToFull(t.name);
                    if (IsValidFilingFolder(alt))
                    {
                        target = alt;
                        break;
                    }
                }
            }

            return target ?? string.Empty;
        }

        private void BindFolderTree()
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)BindFolderTree); return; }
            //RefreshFolderList();
            //RefreshQueues();
            folderTree.BeginUpdate();
            folderTree.Nodes.Clear();

            foreach (var kv in _groups.OrderBy(kv => StripAccountRoot(kv.Key), StringComparer.OrdinalIgnoreCase))
            {
                var list = kv.Value;
                int hi = list.Count(x => x.Confidence >= _autoApproveThreshold);
                int md = list.Count - hi;

                var node = new TreeNode($"{ShortenPath(StripAccountRoot(kv.Key), 48)}  ({list.Count})")
                {
                    Tag = kv.Key,
                    ToolTipText = $"{StripAccountRoot(kv.Key)}   High:{hi}  Med:{md}"
                };
                folderTree.Nodes.Add(node);
            }

            folderTree.EndUpdate();

            // keep previous selection if possible, else select first
            if (!string.IsNullOrEmpty(_selectedGroupKey))
            {
                var match = folderTree.Nodes.Cast<TreeNode>()
                    .FirstOrDefault(n => string.Equals(n.Tag as string, _selectedGroupKey, StringComparison.OrdinalIgnoreCase));
                if (match != null) folderTree.SelectedNode = match;
            }
            if (folderTree.SelectedNode == null && folderTree.Nodes.Count > 0)
                folderTree.SelectedNode = folderTree.Nodes[0];
        }

        private void RebindGridForGroup(string groupKey)
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)(() => RebindGridForGroup(groupKey))); return; }
            List<Services.QueueService.Item> list = null;
            if (!string.IsNullOrEmpty(groupKey))
                _groups.TryGetValue(groupKey, out list);
            list = list ?? new List<Services.QueueService.Item>();

            var view = list.Select(m => new
            {
                m.EntryId,
                m.Subject,
                From = GetPrettyFrom(m.EntryId, m.From),
                Tier = (m.Confidence >= _autoApproveThreshold ? "High" : "Medium"),
                PredictedFull = groupKey,
                Predicted = ShortenPath(StripAccountRoot(groupKey), 60),
                Confidence = m.Confidence.ToString("0.00")
            }).ToList();

            grid.DataSource = view;

            // ensure hidden key column
            if (grid.Columns["EntryId"] == null)
                grid.Columns.Insert(0, new DataGridViewTextBoxColumn { Name = "EntryId", DataPropertyName = "EntryId", Visible = false });
            else
                grid.Columns["EntryId"].Visible = false;

            // per-folder counts + button state
            btnApproveAll.Text = $"Approve Group ({list.Count})";
            btnApproveAll.Enabled = list.Count > 0;

            lblStats.Text =
                $"Folder: {ShortenPath(StripAccountRoot(groupKey), 60)}  " +
                $"High: {list.Count(x => x.Confidence >= _autoApproveThreshold)}  " +
                $"Medium: {list.Count(x => x.Confidence < _autoApproveThreshold)}  " +
                $"Excluded: {_excludedEntryIds.Count}  " +
                $"Total in folder: {list.Count}";
        }

        private async Task ApproveAllInSelectedGroupAsync()
        {
            // Capture BEFORE we mutate queues/tree
            var groupKey = _selectedGroupKey;
            if (string.IsNullOrEmpty(groupKey)) return;

            if (!_groups.TryGetValue(groupKey, out var items) || items == null || items.Count == 0) return;

            // Capture a stable, user-friendly label for the dialog
            var folderCaption = ShortenPath(StripAccountRoot(groupKey), 72);

            int moved = 0;
            foreach (var item in items.ToList())
            {
                string chosen = GetEffectiveTarget(item);
                if (string.IsNullOrEmpty(chosen)) continue;

                await MoveAndLogSelectedAsync(item.EntryId, item.PredictedFolder, chosen, item.Confidence);
                moved++;
            }

            if (_queue != null)
                await _queue.BuildQueuesAsync(_app);

            UI(() => RefreshQueues());

            MessageBox.Show($"Approved {moved} item(s) in \"{folderCaption}\".", "Approve Group");
        }

        // Add this helper method to TaskPaneControl
        private static bool IsValidFilingFolder(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            
            // Exclude system folders (including root Inbox)
            if (FolderMap.IsSystemFolder(path)) return false;
            
            // Allow Inbox subfolders OR user-created folders
            return FolderMap.IsUserCreatedFolder(path);
        }

        // Add this method to refresh folder list and clear stale folder cache
        public void RefreshFolderCache()
        {
            try
            {
                // Clear cached folder list to force refresh from current Outlook state
                _mailFolderPathsFull.Clear();
                
                // Rebuild from current Outlook folder structure
                foreach (var p in FolderMap.GetAllFolderPaths(_app))
                {
                    // Only include valid filing folders (this will exclude system folders and non-existent ones)
                    if (IsValidFilingFolder(p))
                    {
                        // Verify folder still exists before adding
                        try
                        {
                            var ns = _app.Session;
                            var folder = FolderMap.ResolveFolderByPath(p, ns);
                            if (folder != null) // Only add if folder actually exists
                                _mailFolderPathsFull.Add(p);
                        }
                        catch
                        {
                            // Skip folders that can't be resolved (deleted/moved folders)
                        }
                    }
                }

                // Refresh the combo box items
                var items = BuildFolderItems("");
                _suppressFilter = true;
                cmbFolders.Items.Clear();
                foreach (var it in items) cmbFolders.Items.Add(it);
                _suppressFilter = false;
            }
            catch
            {
                // If refresh fails, keep existing folder list
            }
        }

        private async Task BuildQueuesAndRefreshAsync()
        {
            if (_queue == null) return;
            try
            {
                await _queue.BuildQueuesAsync(_app);
                if (this.IsHandleCreated)
                    this.BeginInvoke((Action)(() => RefreshQueues()));
                else
                    RefreshQueues();
            }
            catch
            {
                // Non-fatal: leave Batch empty if a scan fails. You can log if desired.
            }
        }

        // Cclean up training data and refresh everything
        public async Task CleanupAndRefreshAsync()
        {
            await SafeRun(async () =>
            {
                // Build current valid folder set from Outlook, purge stale DB labels
                var validPaths = new List<string>();
                foreach (var p in OutlookClassifierAddIn5.Data.FolderMap.GetAllFolderPaths(_app))
                    validPaths.Add(p);

                var normalized = await _store.NormalizeFolderLabelsAsync(validPaths);
                var cleanedSys = await _store.CleanupSystemFoldersAsync();
                // (optional) _ = await _store.CleanupNonexistentFoldersAsync(validPaths);

                RefreshFolderCache();
                await TrainModelAsync(_modelPath);
                if (_queue != null) await _queue.BuildQueuesAsync(_app);
                RefreshQueues();

                MessageBox.Show($"Normalized {normalized} rows; cleaned system: {cleanedSys}.", "Cleanup Complete");
            });
        }

        // --- Path normalization + display helpers ---

        private static string NormalizePath(string path)
        {
            return string.IsNullOrWhiteSpace(path)
                ? string.Empty
                : path.Replace('\\', '/').Trim().Trim('/');
        }

        private static string StripAccountRoot(string path)
        {
            var p = NormalizePath(path);
            int i = p.IndexOf('/');
            return i >= 0 ? p.Substring(i + 1) : p;
        }

        private static string GetLeaf(string path)
        {
            var p = NormalizePath(path);
            int i = p.LastIndexOf('/');
            return i >= 0 ? p.Substring(i + 1) : p;
        }

        private static string ToShortDisplay(string path)
        {
            // leaf of the path without the account root
            return GetLeaf(StripAccountRoot(path));
        }

        /// <summary>
        /// Map any label (full/short/leaf; any slashes) to a known full Outlook path
        /// from the current cache. Returns empty if ambiguous or not found.
        /// </summary>
        private string CanonicalizeToFull(string label)
        {
            var s = NormalizePath(label);
            if (string.IsNullOrEmpty(s)) return string.Empty;

            // Exact path match?
            var exact = _mailFolderPathsFull.FirstOrDefault(p =>
                string.Equals(NormalizePath(p), s, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(exact)) return exact;

            // Unique leaf match?
            var leafMatches = _mailFolderPathsFull.Where(p =>
                string.Equals(GetLeaf(p), s, StringComparison.OrdinalIgnoreCase)).ToList();
            if (leafMatches.Count == 1) return leafMatches[0];

            // Unique "display" (without account root)?
            var dispMatches = _mailFolderPathsFull.Where(p =>
                string.Equals(StripAccountRoot(p), s, StringComparison.OrdinalIgnoreCase)).ToList();
            if (dispMatches.Count == 1) return dispMatches[0];

            // Ambiguous or unknown -> don't guess
            return string.Empty;
        }

        private void ClampComboDropDown()
        {
            try
            {
                // Available space below the ComboBox
                var screen = Screen.FromControl(cmbFolders).WorkingArea;
                var belowPt = cmbFolders.PointToScreen(new System.Drawing.Point(0, cmbFolders.Height));
                int spaceBelow = screen.Bottom - belowPt.Y - 8; // padding

                // Item height heuristic
                int itemH = cmbFolders.ItemHeight > 0 ? cmbFolders.ItemHeight : cmbFolders.Font.Height + 6;

                // Limit visible items to what fits on screen, minimum 4
                int maxItemsBelow = Math.Max(4, spaceBelow / Math.Max(1, itemH));
                cmbFolders.MaxDropDownItems = Math.Min(cmbFolders.MaxDropDownItems, maxItemsBelow);

                // Hard cap to avoid super tall lists
                cmbFolders.DropDownHeight = Math.Min(spaceBelow, 400);
                cmbFolders.IntegralHeight = false; // ensure height is honored

                // Width: fit longest item but cap it (and not smaller than control)
                int desiredWidth = ComputeDropDownWidth(cmbFolders, 600);
                int screenCap = screen.Width / 2; // don’t let it be half the screen
                cmbFolders.DropDownWidth = Math.Min(desiredWidth, screenCap);
            }
            catch { /* layout races: ignore */ }
        }

        private static int ComputeDropDownWidth(ComboBox cb, int hardCap)
        {
            int width = cb.Width;
            for (int i = 0; i < cb.Items.Count; i++)
            {
                var text = cb.GetItemText(cb.Items[i]) ?? string.Empty;
                var sz = TextRenderer.MeasureText(text, cb.Font);
                if (sz.Width > width) width = sz.Width;
            }
            // room for scrollbar + padding
            width += SystemInformation.VerticalScrollBarWidth + 12;
            return Math.Min(width, hardCap);
        }

        private static string GetSenderSmtpAddress(Outlook.MailItem mail)
        {
            if (mail == null) return string.Empty;

            try
            {
                // If Outlook already gives SMTP, use it
                if (!string.IsNullOrEmpty(mail.SenderEmailAddress) &&
                    !mail.SenderEmailAddress.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                {
                    return mail.SenderEmailAddress;
                }

                Outlook.AddressEntry sender = mail.Sender;
                if (sender == null)
                {
                    try { mail.Recipients?.ResolveAll(); } catch { }
                    sender = mail.Sender;
                }
                if (sender == null) return mail.SenderEmailAddress ?? string.Empty;

                // 1) Exchange users (internal or remote)
                if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var ex = sender.GetExchangeUser();
                    if (ex != null && !string.IsNullOrEmpty(ex.PrimarySmtpAddress))
                        return ex.PrimarySmtpAddress;
                }

                // 2) SMTP entries (external)
                if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olSmtpAddressEntry)
                {
                    // Often Address is already SMTP here
                    if (!string.IsNullOrEmpty(sender.Address) &&
                        !sender.Address.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                        return sender.Address;
                }

                // 3) Fallback: PR_SMTP_ADDRESS on the AddressEntry
                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                try
                {
                    var pa = sender.PropertyAccessor;
                    var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                    if (!string.IsNullOrEmpty(smtp)) return smtp;
                }
                catch { /* some stores may not expose it */ }

                // 4) Last-resort fallbacks
                if (!string.IsNullOrEmpty(mail.SenderEmailAddress) &&
                    !mail.SenderEmailAddress.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                    return mail.SenderEmailAddress;

                // Give at least a name if all else fails
                return sender.Name ?? mail.SenderName ?? (mail.ReplyRecipients?.Count > 0 ? mail.ReplyRecipients[1].Address : string.Empty) ?? string.Empty;
            }
            catch
            {
                return mail.SenderEmailAddress ?? string.Empty;
            }
        }

        private static string GetSenderDisplayName(Outlook.MailItem mail)
        {
            try
            {
                var sender = mail?.Sender;
                if (sender == null) return mail?.SenderName ?? string.Empty;

                // Prefer Exchange display name if available
                if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var ex = sender.GetExchangeUser();
                    if (ex != null && !string.IsNullOrEmpty(ex.Name))
                        return ex.Name;
                }
                return sender.Name ?? mail?.SenderName ?? string.Empty;
            }
            catch { return mail?.SenderName ?? string.Empty; }
        }

        public void BindFromCache(OutlookClassifierAddIn5.Services.QueueCache cache)
        {
            if (this.InvokeRequired) { this.BeginInvoke((System.Action)(() => BindFromCache(cache))); return; }
            if (cache == null) return;

            // Fill the queue from cache and refresh the UI
            _queue.SetFromCache(cache);
            RefreshQueues();

            // Optional: tell the user we loaded cached results
            if (lblStats != null)
                lblStats.Text = "Loaded cached results • Refreshing…";
        }

        // inside OutlookClassifierAddIn5.UI.TaskPaneControl
        public void Deactivate()
        {
            try { _showCts?.Cancel(); } catch { }
        }

        private static bool LooksLegacyDn(string s)
        {
            return !string.IsNullOrEmpty(s) &&
                   s.StartsWith("/O=", StringComparison.OrdinalIgnoreCase);
        }

        private string GetPrettyFrom(string entryId, string raw)
        {
            // If we’ve already resolved it, reuse
            string cached;
            if (!string.IsNullOrEmpty(entryId) && _prettyFromCache.TryGetValue(entryId, out cached))
                return cached;

            // Already SMTP-looking? Just show it.
            if (!LooksLegacyDn(raw))
                return raw ?? string.Empty;

            // Resolve via MailItem (only for visible rows, so this is cheap)
            Outlook.MailItem mi = null;
            try
            {
                var ns = _app?.Session;
                mi = ns?.GetItemFromID(entryId) as Outlook.MailItem;
                if (mi == null) return raw ?? string.Empty;

                string display = mi.SenderName ?? string.Empty;
                string smtp = mi.SenderEmailAddress ?? string.Empty;

                try
                {
                    // If SMTP is still legacy, try ExchangeUser + PR_SMTP_ADDRESS
                    if (LooksLegacyDn(smtp))
                    {
                        var sender = mi.Sender;
                        if (sender != null)
                        {
                            if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                                sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                            {
                                var ex = sender.GetExchangeUser();
                                if (ex != null)
                                {
                                    if (!string.IsNullOrEmpty(ex.PrimarySmtpAddress)) smtp = ex.PrimarySmtpAddress;
                                    if (!string.IsNullOrEmpty(ex.Name)) display = ex.Name;
                                }
                            }

                            if (LooksLegacyDn(smtp))
                            {
                                const string PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                                try
                                {
                                    var pa = sender.PropertyAccessor;
                                    var addr = pa.GetProperty(PR_SMTP) as string;
                                    if (!string.IsNullOrEmpty(addr)) smtp = addr;
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { }

                var pretty = !string.IsNullOrEmpty(smtp)
                    ? (string.IsNullOrEmpty(display) ? smtp : (display + " <" + smtp + ">"))
                    : (string.IsNullOrEmpty(display) ? raw ?? string.Empty : display);

                if (!string.IsNullOrEmpty(entryId))
                    _prettyFromCache[entryId] = pretty;

                return pretty;
            }
            catch
            {
                return raw ?? string.Empty;
            }
            finally
            {
                if (mi != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(mi);
            }
        }


    }
}