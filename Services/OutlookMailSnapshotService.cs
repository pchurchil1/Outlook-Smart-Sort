using System;
using System.Security.Cryptography;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Services
{
    public sealed class OutlookMailSnapshotService
    {
        private const string PrInternetMessageId = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        private readonly Outlook.Application _app;

        public OutlookMailSnapshotService(Outlook.Application app)
        {
            _app = app;
        }

        public MailFeatureDto GetSnapshotByEntryId(string entryId, int bodyLimit)
        {
            if (string.IsNullOrEmpty(entryId)) return null;
            var ns = _app.Session;
            var mail = ns.GetItemFromID(entryId) as Outlook.MailItem;
            return mail == null ? null : CreateSnapshot(mail, bodyLimit);
        }

        public static MailFeatureDto CreateSnapshot(Outlook.MailItem mail, int bodyLimit)
        {
            if (mail == null) return null;

            var parent = mail.Parent as Outlook.MAPIFolder;
            var currentPath = GetFolderPath(parent);
            var rawSender = SafeString(() => mail.SenderEmailAddress);
            var smtp = SenderResolutionService.GetSenderSmtpAddress(mail);
            var featureSender = SenderResolutionService.ChooseFeatureAddress(smtp, rawSender);
            var receivedUtc = SafeDate(() => mail.ReceivedTime);
            var body = SafeString(() => mail.Body);

            if (bodyLimit >= 0 && body.Length > bodyLimit)
                body = body.Substring(0, bodyLimit);

            return new MailFeatureDto
            {
                EntryId = SafeString(() => mail.EntryID),
                StoreId = SafeString(() => parent == null ? string.Empty : parent.StoreID),
                Subject = SafeString(() => mail.Subject),
                BodySnippet = body,
                FromAddress = featureSender,
                SenderDomain = SenderResolutionService.ExtractDomain(featureSender),
                HasAttachments = SafeBool(() => mail.Attachments != null && mail.Attachments.Count > 0),
                ReceivedUtc = receivedUtc.HasValue ? receivedUtc.Value.ToUniversalTime() : (DateTime?)null,
                CurrentFolderPath = currentPath,
                ConversationId = SafeString(() => mail.ConversationID),
                InternetMessageId = SafeProperty(mail, PrInternetMessageId)
            };
        }

        public static string GetFolderPath(Outlook.MAPIFolder folder)
        {
            if (folder == null) return string.Empty;

            try
            {
                if (!string.IsNullOrEmpty(folder.FolderPath))
                    return FolderPathNormalizer.Normalize(folder.FolderPath);
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Could not read FolderPath: " + ex.Message);
            }

            try
            {
                var path = folder.Name ?? string.Empty;
                var parent = folder.Parent as Outlook.MAPIFolder;
                while (parent != null)
                {
                    path = (parent.Name ?? string.Empty) + "/" + path;
                    parent = parent.Parent as Outlook.MAPIFolder;
                }

                return FolderPathNormalizer.Normalize(path);
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Could not build folder path: " + ex.Message);
                return string.Empty;
            }
        }

        public static string NormalizedSubjectHash(string subject)
        {
            var normalized = (subject ?? string.Empty).Trim().ToLowerInvariant();
            while (normalized.StartsWith("re:", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("fw:", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("fwd:", StringComparison.OrdinalIgnoreCase))
            {
                var colon = normalized.IndexOf(':');
                normalized = colon >= 0 ? normalized.Substring(colon + 1).Trim() : normalized.Trim();
            }

            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized));
                return BitConverter.ToString(bytes).Replace("-", string.Empty);
            }
        }

        private static string SafeString(Func<string> get)
        {
            try { return get() ?? string.Empty; }
            catch (Exception ex) { AppLogger.Warn("Outlook string read failed: " + ex.Message); return string.Empty; }
        }

        private static bool SafeBool(Func<bool> get)
        {
            try { return get(); }
            catch (Exception ex) { AppLogger.Warn("Outlook bool read failed: " + ex.Message); return false; }
        }

        private static DateTime? SafeDate(Func<DateTime> get)
        {
            try
            {
                var value = get();
                return value == DateTime.MinValue ? (DateTime?)null : value;
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Outlook date read failed: " + ex.Message);
                return null;
            }
        }

        private static string SafeProperty(Outlook.MailItem mail, string propertyUri)
        {
            try
            {
                var pa = mail.PropertyAccessor;
                return pa.GetProperty(propertyUri) as string ?? string.Empty;
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Outlook property read failed: " + propertyUri + " " + ex.Message);
                return string.Empty;
            }
        }
    }
}
