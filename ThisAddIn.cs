using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookClassifierAddIn5.Data;
using OutlookClassifierAddIn5.ML;
using OutlookClassifierAddIn5.Services;
using OutlookClassifierAddIn5.UI;

namespace OutlookClassifierAddIn5
{
    public partial class ThisAddIn
    {
        private CancellationTokenSource _queueBuildCts;
        private CustomTaskPane _pane;
        private Outlook.Explorer _explorer;
        private bool _paneInitialized;
        private bool _visibleChangedHooked;
        private bool _selectionHooked;
        private bool _initialActivationComplete;

        internal TaskPaneControl PaneControl { get; private set; }

        internal static ModelService ModelSvc;
        internal static FeedbackStore Store;
        internal static ScanService Scanner;
        internal static QueueService Queue;

        private string _modelPath;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _modelPath = System.IO.Path.Combine(AppLogger.DataDirectory, "model.zip");
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_modelPath));
            AppLogger.Info("Outlook Smart Sort startup initialized.");
        }

        private void EnsurePane()
        {
            if (_pane != null) return;

            EnsureServices();

            PaneControl = new TaskPaneControl(Application, ModelSvc, Store, Queue, _modelPath);
            _pane = CustomTaskPanes.Add(PaneControl, "Email Classifier");
            if (_pane.Width < 900) _pane.Width = 900;
            _pane.Visible = false;

            if (!_visibleChangedHooked)
            {
                _pane.VisibleChanged += Pane_VisibleChanged;
                _visibleChangedHooked = true;
            }

            AppLogger.Info("Task pane created.");
        }

        private void EnsureServices()
        {
            if (Store == null)
            {
                Store = new FeedbackStore();
                Store.Initialize();
                AppLogger.Info("Feedback store initialized.");
            }

            if (ModelSvc == null)
            {
                ModelSvc = new ModelService();
                if (System.IO.File.Exists(_modelPath))
                {
                    if (!ModelSvc.TryLoad(_modelPath))
                        AppLogger.Warn("Existing model could not be loaded; retraining will be required.");
                }
                else
                {
                    AppLogger.Warn("No model file found. User can retrain after seed data is available.");
                }
            }

            if (Scanner == null)
                Scanner = new ScanService(Application, Store);

            if (Queue == null)
                Queue = new QueueService(ModelSvc, Store);
        }

        private void Pane_VisibleChanged(object sender, EventArgs e)
        {
            if (_pane == null || PaneControl == null) return;

            if (_pane.Visible)
            {
                if (!_paneInitialized)
                {
                    PaneControl.Init();
                    _paneInitialized = true;
                    AppLogger.Info("Task pane initialized.");
                }

                HookSelectionChanged();
                StartPaneActivationAsync();
                Explorer_SelectionChange();
            }
            else
            {
                CancelQueueBuild("Pane hidden.");
                PaneControl.Deactivate();
                UnhookSelectionChanged();
            }
        }

        private async void StartPaneActivationAsync()
        {
            CancelQueueBuild("Restarting pane activation.");
            _queueBuildCts = new CancellationTokenSource();
            var ct = _queueBuildCts.Token;

            try
            {
                if (!_initialActivationComplete && !ModelSvc.IsLoaded)
                {
                    var cleanupVersion = await Store.GetMetaValueAsync("system_folders_cleanup_v1");
                    if (string.IsNullOrEmpty(cleanupVersion))
                    {
                        var cleaned = await Store.CleanupSystemFoldersAsync();
                        await Store.SetMetaValueAsync("system_folders_cleanup_v1", "completed");
                        AppLogger.Info("Initial system-folder cleanup removed " + cleaned + " rows.");
                    }

                    PaneControl.SetStatus("No model trained yet. Scanning folders...");
                    var bodyLength = await Store.GetBodySnippetLengthAsync();
                    await Scanner.EnsureSeedTrainingDataAsync(new ScanService.ScanOptions
                    {
                        BodySamplePerFolder = bodyLength > 0 ? 20 : 0,
                        BodySnippetLength = bodyLength
                    }, ct);

                    ct.ThrowIfCancellationRequested();
                    PaneControl.SetStatus("Training model...");
                    await PaneControl.TrainModelAsync(_modelPath);
                    _initialActivationComplete = true;
                }
                else
                {
                    _initialActivationComplete = true;
                    PaneControl.SetStatus(ModelSvc.IsLoaded ? "Loading model..." : "No model trained yet. Click Retrain.");
                }

                if (!ModelSvc.IsLoaded)
                {
                    PaneControl.SetStatus("No model trained yet. Click Retrain.");
                    return;
                }

                ct.ThrowIfCancellationRequested();
                PaneControl.SetStatus("Building suggestions...");
                await Queue.BuildQueuesAsync(Application, 90, ct);

                if (!ct.IsCancellationRequested && PaneControl != null)
                    PaneControl.RefreshQueues();

                Explorer_SelectionChange();
            }
            catch (OperationCanceledException)
            {
                AppLogger.Info("Pane activation cancelled.");
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Pane activation failed.");
                if (PaneControl != null) PaneControl.SetStatus("Suggestion build failed. See local log.");
            }
        }

        private void Explorer_SelectionChange()
        {
            try
            {
                if (_pane == null || !_pane.Visible || PaneControl == null) return;

                var explorer = _explorer ?? Application.ActiveExplorer();
                var sel = explorer != null ? explorer.Selection : null;
                if (sel == null || sel.Count <= 0) return;

                object obj = sel[1];
                var mail = obj as Outlook.MailItem;
                if (mail == null) return;

                PaneControl.BeginShowSelectedAsync(mail.EntryID);
                PaneControl.SelectSingleTab();
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Selection change handling failed: " + ex.Message);
            }
        }

        private void HookSelectionChanged()
        {
            if (_selectionHooked) return;

            _explorer = Application.ActiveExplorer();
            if (_explorer == null) return;

            _explorer.SelectionChange += Explorer_SelectionChange;
            _selectionHooked = true;
            AppLogger.Info("Explorer selection handler hooked.");
        }

        private void UnhookSelectionChanged()
        {
            if (!_selectionHooked) return;

            try
            {
                if (_explorer != null)
                    _explorer.SelectionChange -= Explorer_SelectionChange;
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Explorer selection unhook failed: " + ex.Message);
            }
            finally
            {
                _selectionHooked = false;
            }
        }

        private void CancelQueueBuild(string reason)
        {
            if (_queueBuildCts == null) return;
            try
            {
                _queueBuildCts.Cancel();
                AppLogger.Info("Queue build cancelled. " + reason);
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Queue cancellation failed: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                CancelQueueBuild("Shutdown.");
                UnhookSelectionChanged();

                if (_pane != null)
                {
                    if (_visibleChangedHooked)
                    {
                        _pane.VisibleChanged -= Pane_VisibleChanged;
                        _visibleChangedHooked = false;
                    }

                    CustomTaskPanes.Remove(_pane);
                    _pane = null;
                }

                AppLogger.Info("Outlook Smart Sort shutdown complete.");
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Shutdown cleanup failed.");
            }
        }

        internal void TogglePane()
        {
            EnsurePane();
            _pane.Visible = !_pane.Visible;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
