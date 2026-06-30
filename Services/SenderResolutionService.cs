using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Services
{
    public static class SenderResolutionService
    {
        private const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public static bool LooksLegacyDn(string value)
        {
            return !string.IsNullOrEmpty(value) &&
                   value.StartsWith("/O=", StringComparison.OrdinalIgnoreCase);
        }

        public static string ExtractDomain(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return string.Empty;
            var at = address.IndexOf('@');
            if (at < 0 || at == address.Length - 1) return string.Empty;
            return address.Substring(at + 1).Trim().ToLowerInvariant();
        }

        public static string ChooseFeatureAddress(string resolvedSmtp, string fallback)
        {
            if (!string.IsNullOrWhiteSpace(resolvedSmtp) && !LooksLegacyDn(resolvedSmtp))
                return resolvedSmtp.Trim();

            if (!string.IsNullOrWhiteSpace(fallback) && !LooksLegacyDn(fallback))
                return fallback.Trim();

            return string.Empty;
        }

        public static string ComposePretty(string display, string smtpOrRaw)
        {
            var smtp = smtpOrRaw ?? string.Empty;
            var name = display ?? string.Empty;

            if (!string.IsNullOrEmpty(smtp) && !LooksLegacyDn(smtp))
                return string.IsNullOrEmpty(name) ? smtp : name + " <" + smtp + ">";

            return name.Length > 0 ? name : smtp;
        }

        public static string GetSenderSmtpAddress(Outlook.MailItem mail)
        {
            if (mail == null) return string.Empty;

            try
            {
                var raw = mail.SenderEmailAddress ?? string.Empty;
                if (!string.IsNullOrEmpty(raw) && !LooksLegacyDn(raw))
                    return raw;

                var sender = mail.Sender;
                if (sender == null)
                {
                    try { if (mail.Recipients != null) mail.Recipients.ResolveAll(); } catch (Exception ex) { AppLogger.Warn("Could not resolve recipients while resolving sender SMTP: " + ex.Message); }
                    sender = mail.Sender;
                }

                if (sender == null) return raw;

                if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exUser = sender.GetExchangeUser();
                    if (exUser != null && !string.IsNullOrEmpty(exUser.PrimarySmtpAddress))
                        return exUser.PrimarySmtpAddress;
                }

                if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olSmtpAddressEntry)
                {
                    if (!string.IsNullOrEmpty(sender.Address) && !LooksLegacyDn(sender.Address))
                        return sender.Address;
                }

                try
                {
                    var pa = sender.PropertyAccessor;
                    var smtp = pa.GetProperty(PrSmtpAddress) as string;
                    if (!string.IsNullOrEmpty(smtp) && !LooksLegacyDn(smtp))
                        return smtp;
                }
                catch (Exception ex)
                {
                    AppLogger.Warn("Could not read sender SMTP property: " + ex.Message);
                }

                return LooksLegacyDn(raw) ? string.Empty : raw;
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Failed resolving sender SMTP address.");
                return mail.SenderEmailAddress ?? string.Empty;
            }
        }

        public static string GetSenderDisplayName(Outlook.MailItem mail)
        {
            try
            {
                var sender = mail == null ? null : mail.Sender;
                if (sender == null) return mail == null ? string.Empty : (mail.SenderName ?? string.Empty);

                if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exUser = sender.GetExchangeUser();
                    if (exUser != null && !string.IsNullOrEmpty(exUser.Name))
                        return exUser.Name;
                }

                return sender.Name ?? (mail == null ? string.Empty : (mail.SenderName ?? string.Empty));
            }
            catch (Exception ex)
            {
                AppLogger.Warn("Could not resolve sender display name: " + ex.Message);
                return mail == null ? string.Empty : (mail.SenderName ?? string.Empty);
            }
        }
    }
}
