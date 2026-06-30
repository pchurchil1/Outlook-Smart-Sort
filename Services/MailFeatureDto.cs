using System;

namespace OutlookClassifierAddIn5.Services
{
    public sealed class MailFeatureDto
    {
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public string Subject { get; set; }
        public string BodySnippet { get; set; }
        public string FromAddress { get; set; }
        public string SenderDomain { get; set; }
        public bool HasAttachments { get; set; }
        public DateTime? ReceivedUtc { get; set; }
        public string CurrentFolderPath { get; set; }
        public string ConversationId { get; set; }
        public string InternetMessageId { get; set; }
    }
}
