using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using OutlookClassifierAddIn5.UI;
using OutlookClassifierAddIn5.ML;
using OutlookClassifierAddIn5.Data;
using OutlookClassifierAddIn5.Services;

namespace OutlookClassifierAddIn5
{
    public partial class ThisAddIn
    {
        private CancellationTokenSource _queueBuildCts;
        private CustomTaskPane _pane;
        private Outlook.Explorer _explorer;
        private bool _paneInitialized;
        internal TaskPaneControl PaneControl { get; private set; }

        internal static ModelService ModelSvc;
        internal static FeedbackStore Store;
        internal static ScanService Scanner;
        internal static QueueService Queue;

        private string _modelPath;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _modelPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookClassifier", "model.zip");
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_modelPath));

            Store = new FeedbackStore();
            Store.Initialize();

            ModelSvc = new ModelService();
            if (System.IO.File.Exists(_modelPath))
            {
                try { ModelSvc.Load(_modelPath); } catch { /* ignore — will retrain */ }
            }

            Scanner = new ScanService(Application, Store);
            Queue = new QueueService(ModelSvc, Store);

            PaneControl = new TaskPaneControl(Application, ModelSvc, Store, Queue, _modelPath);
            _pane = this.CustomTaskPanes.Add(PaneControl, "Email Classifier");
            if (_pane.Width < 900) _pane.Width = 900;
            _pane.Visible = false;

            _explorer = Application.ActiveExplorer();
            if (_explorer != null)
                _explorer.SelectionChange += Explorer_SelectionChange;

            // Load whatever is currently selected (if any)
            Explorer_SelectionChange();
        }

        private void EnsurePane()
        {
            if (_pane != null) return;

            _modelPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookClassifier", "model.zip");
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_modelPath));

            // Create services now (on demand)
            Store = new FeedbackStore();
            Store.Initialize();

            ModelSvc = new ModelService();
            if (System.IO.File.Exists(_modelPath))
            {
                try { ModelSvc.Load(_modelPath); } catch { /* ok, will retrain when shown */ }
            }

            Scanner = new ScanService(Application, Store);
            Queue = new QueueService(ModelSvc, Store);

            PaneControl = new TaskPaneControl(Application, ModelSvc, Store, Queue, _modelPath);
            _pane = this.CustomTaskPanes.Add(PaneControl, "Email Classifier");
            if (_pane.Width < 900) _pane.Width = 900;
            _pane.Visible = false;

            _pane.VisibleChanged += Pane_VisibleChanged;
        }

        private void Pane_VisibleChanged(object sender, System.EventArgs e)
        {
            if (_pane == null || PaneControl == null) return;

            if (_pane.Visible)
            {
                if (!_paneInitialized)
                {
                    _paneInitialized = true;

                    // Light UI -> heavy classifier init now
                    PaneControl.Init();

                    // Hook explorer selection only while active
                    _explorer = Application.ActiveExplorer();
                    if (_explorer != null)
                        _explorer.SelectionChange += Explorer_SelectionChange;

                    // Kick off cleanup/scan/train AFTER activation
                    PaneControl.SafeRun(async () =>
                    {
                        // one-time cleanup
                        var cleanupVersion = await Store.GetMetaValueAsync("system_folders_cleanup_v1");
                        if (string.IsNullOrEmpty(cleanupVersion))
                        {
                            var cleanedCount = await Store.CleanupSystemFoldersAsync();
                            await Store.SetMetaValueAsync("system_folders_cleanup_v1", "completed");
                            if (cleanedCount > 0)
                                System.Windows.Forms.MessageBox.Show($"Cleaned up {cleanedCount} system folder entries from training data.", "Training Data Cleaned");
                        }

                        await Scanner.EnsureSeedTrainingDataAsync();
                        await PaneControl.TrainModelAsync(savePath: _modelPath);
                        await Queue.BuildQueuesAsync(Application);
                        var pane = PaneControl;
                        if (pane != null)
                        {
                            if (pane.IsHandleCreated)
                                pane.BeginInvoke((System.Action)(() => pane.RefreshQueues()));
                            else
                                pane.RefreshQueues(); // ok if handle not created yet (rare here)
                        }

                        // preload currently selected mail (optional)
                        Explorer_SelectionChange();
                    });
                }
            }
            else
            {
                // Pane hidden: unhook selection to reduce churn
                try { if (_explorer != null) _explorer.SelectionChange -= Explorer_SelectionChange; } catch { }
            }
        }


        private void Explorer_SelectionChange()
        {
            try
            {
                var sel = _explorer != null ? _explorer.Selection : null;
                if (sel != null && sel.Count > 0)
                {
                    object obj = sel[1];
                    var mail = obj as Outlook.MailItem;
                    if (mail != null)
                    {
                        // fire-and-forget; pane handles cancellation/debounce
                        PaneControl.BeginShowSelectedAsync(mail.EntryID);
                        // optionally flip to the Single tab:
                        PaneControl.SelectSingleTab();
                    }
                }
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (_explorer != null)
                    _explorer.SelectionChange -= Explorer_SelectionChange;

                if (_pane != null)
                {
                    this.CustomTaskPanes.Remove(_pane);
                    _pane = null;
                }
            }
            catch { }
        }

        internal async void TogglePane()
        {
            EnsurePane();

            if (!_pane.Visible)
            {
                _pane.Visible = true;

                if (!_paneInitialized)
                {
                    _paneInitialized = true;
                    PaneControl.Init();
                }

                // Cancel any prior build
                if (_queueBuildCts != null) { try { _queueBuildCts.Cancel(); } catch { } }
                _queueBuildCts = new CancellationTokenSource();
                var ct = _queueBuildCts.Token;

                // 1) FAST PASS (e.g., 45 days) — background, then bind
                await Task.Run(async () =>
                {
                    try
                    {
                        await Queue.BuildQueuesAsync(this.Application, 45, ct);
                    }
                    catch { }
                }, ct);

                // Bind on UI
                if (!ct.IsCancellationRequested)
                {
                    var pane = PaneControl;
                    if (pane != null)
                    {
                        if (pane.IsHandleCreated)
                            pane.BeginInvoke((System.Action)(() => pane.RefreshQueues()));
                        else
                            pane.RefreshQueues();
                    }
                }

                // 2) EXTEND PASS (e.g., 120 days) — background, then bind + save cache
                if (!ct.IsCancellationRequested)
                {
                    await Task.Run(async () =>
                    {
                        try
                        {
                            await Queue.BuildQueuesAsync(this.Application, 120, ct);
                            // optional: save cache here if you added it earlier
                            var folders = FolderMap.GetAllFolderPaths(this.Application) ?? new System.Collections.Generic.List<string>();
                            var cache = Queue.SnapshotToCache(folders);
                            var cachePath = System.IO.Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                "OutlookClassifier", "queuecache.json");
                            Queue.SaveCache(cachePath, cache);
                        }
                        catch { }
                    }, ct);

                    if (!ct.IsCancellationRequested)
                    {
                        var pane2 = PaneControl;
                        if (pane2 != null)
                        {
                            if (pane2.IsHandleCreated)
                                pane2.BeginInvoke((System.Action)(() => pane2.RefreshQueues()));
                            else
                                pane2.RefreshQueues();
                        }
                    }
                }
            }
            else
            {
                _pane.Visible = false;
                PaneControl.Deactivate();

                // Cancel any ongoing builds when the pane hides
                if (_queueBuildCts != null) { try { _queueBuildCts.Cancel(); } catch { } }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
