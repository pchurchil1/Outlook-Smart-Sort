using System;
using OutlookClassifierAddIn5.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Services
{
    public sealed class OutlookMoveService
    {
        private readonly Outlook.Application _app;
        private readonly int _bodySnippetLimit;

        public OutlookMoveService(Outlook.Application app, int bodySnippetLimit)
        {
            _app = app;
            _bodySnippetLimit = bodySnippetLimit;
        }

        public MoveResult MoveToFolder(string entryId, string predictedFolderPath, string chosenFolderPath, double confidence)
        {
            var chosen = FolderPathNormalizer.Normalize(chosenFolderPath);
            if (string.IsNullOrEmpty(chosen))
                throw new InvalidOperationException("Choose a valid filing folder.");

            if (FolderMap.IsSystemFolder(chosen))
                throw new InvalidOperationException("System folders cannot be used as filing targets.");

            var ns = _app.Session;
            var mail = ns.GetItemFromID(entryId) as Outlook.MailItem;
            if (mail == null)
                throw new InvalidOperationException("Could not locate the message.");

            var snapshot = OutlookMailSnapshotService.CreateSnapshot(mail, _bodySnippetLimit);
            var sourcePath = snapshot == null ? string.Empty : snapshot.CurrentFolderPath;
            var target = FolderMap.ResolveFolderByPath(chosen, ns);

            if (target == null)
                throw new InvalidOperationException("The target folder does not exist.");

            try
            {
                if (target.DefaultItemType != Outlook.OlItemType.olMailItem)
                    throw new InvalidOperationException("The target folder is not a mail folder.");
            }
            catch (InvalidOperationException) { throw; }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Could not validate target folder type.");
                throw new InvalidOperationException("The target folder could not be validated.");
            }

            var moved = mail.Move(target) as Outlook.MailItem;
            var newId = moved == null ? entryId : (moved.EntryID ?? entryId);
            var storeId = snapshot == null ? string.Empty : snapshot.StoreId;
            if (moved != null)
            {
                try
                {
                    var movedParent = moved.Parent as Outlook.MAPIFolder;
                    if (movedParent != null && !string.IsNullOrEmpty(movedParent.StoreID))
                        storeId = movedParent.StoreID;
                }
                catch (Exception ex)
                {
                    AppLogger.Warn("Could not read moved item store id: " + ex.Message);
                }
            }

            return new MoveResult
            {
                OldEntryId = entryId ?? string.Empty,
                NewEntryId = newId ?? string.Empty,
                StoreId = storeId ?? string.Empty,
                SourcePath = sourcePath ?? string.Empty,
                DestPath = chosen,
                PredictedPath = FolderPathNormalizer.Normalize(predictedFolderPath),
                Confidence = confidence,
                FeatureBeforeMove = snapshot
            };
        }

        public void MoveBack(string newEntryId, string oldEntryId, string sourceFolderPath)
        {
            var ns = _app.Session;
            Outlook.MailItem item = null;
            try { item = ns.GetItemFromID(newEntryId) as Outlook.MailItem; } catch (Exception ex) { AppLogger.Warn("Undo lookup by new EntryId failed: " + ex.Message); }
            if (item == null)
            {
                try { item = ns.GetItemFromID(oldEntryId) as Outlook.MailItem; } catch (Exception ex) { AppLogger.Warn("Undo lookup by old EntryId failed: " + ex.Message); }
            }

            if (item == null)
                throw new InvalidOperationException("Could not find the message to undo.");

            var source = FolderMap.ResolveFolderByPath(FolderPathNormalizer.Normalize(sourceFolderPath), ns);
            if (source == null)
                throw new InvalidOperationException("Could not find the original folder.");

            item.Move(source);
        }
    }

    public sealed class MoveResult
    {
        public string OldEntryId { get; set; }
        public string NewEntryId { get; set; }
        public string StoreId { get; set; }
        public string SourcePath { get; set; }
        public string DestPath { get; set; }
        public string PredictedPath { get; set; }
        public double Confidence { get; set; }
        public MailFeatureDto FeatureBeforeMove { get; set; }
    }
}
