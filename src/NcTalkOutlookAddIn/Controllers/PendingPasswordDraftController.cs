// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // A saved password draft is the durable pending record. It is dispatched
    // only after the matching primary item appears as Sent in its exact folder.
    internal sealed class PendingPasswordDraftController : IDisposable
    {
        private const string AttemptProperty = "X-NCC-SEND-ATTEMPT";
        private const string StateProperty = "X-NCC-PASSWORD-STATE";
        private const string PayloadProperty = "X-NCC-PASSWORD-PAYLOAD";
        private const string SentFolderProperty = "X-NCC-SENT-FOLDER";
        private const string SentStoreProperty = "X-NCC-SENT-STORE";
        private const string ArmedState = "armed";
        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes("NCC.PendingPasswordDraft.v1");

        private readonly NextcloudTalkAddIn _owner;
        private readonly ComposeShareLifecycleController _compose;
        private readonly Dictionary<string, FolderWatch> _folderWatches =
            new Dictionary<string, FolderWatch>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<DraftReference>> _drafts =
            new Dictionary<string, List<DraftReference>>(
                StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _dispatching =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool _disposed;

        internal PendingPasswordDraftController(
            NextcloudTalkAddIn owner,
            ComposeShareLifecycleController compose)
        {
            _owner = owner;
            _compose = compose;
        }

        internal void Start()
        {
            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;
            var attempts = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            try
            {
                session = _owner.OutlookApplication != null
                    ? _owner.OutlookApplication.Session
                    : null;
                stores = session != null ? session.Stores : null;
                int count = stores != null ? stores.Count : 0;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Store store = null;
                    Outlook.MAPIFolder drafts = null;
                    try
                    {
                        store = stores[i];
                        drafts = store != null
                            ? store.GetDefaultFolder(
                                Outlook.OlDefaultFolders.olFolderDrafts)
                            : null;
                        ScanDraftFolder(drafts, session, attempts);
                    }
                    finally
                    {
                        Release(drafts, "recovery Drafts folder");
                        Release(store, "recovery Store");
                    }
                }
                foreach (string attempt in attempts)
                {
                    if (IsPrimarySent(session, attempt))
                    {
                        DispatchAttempt(attempt);
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Pending password draft recovery failed; drafts retained.",
                    ex);
            }
            finally
            {
                Release(stores, "recovery Stores");
                Release(session, "recovery Session");
            }
        }

        internal bool TryArm(
            Outlook.MailItem primary,
            string composeKey,
            List<SeparatePasswordDispatchEntry> queue)
        {
            if (queue == null || queue.Count == 0)
            {
                return true;
            }

            string attempt = Guid.NewGuid().ToString("N");
            Outlook.MAPIFolder sentFolder = null;
            var created = new List<Outlook.MailItem>();
            try
            {
                if (primary == null || primary.DeleteAfterSubmit)
                {
                    throw new InvalidOperationException(
                        "The primary message has no verifiable sent copy.");
                }
                sentFolder = ResolveExactSentFolder(primary);
                string sentFolderId = sentFolder.EntryID;
                string sentStoreId = sentFolder.StoreID;
                if (string.IsNullOrWhiteSpace(sentFolderId)
                    || string.IsNullOrWhiteSpace(sentStoreId))
                {
                    throw new InvalidOperationException(
                        "The exact Sent folder could not be identified.");
                }

                List<SeparatePasswordDispatchEntry> expanded =
                    _compose.ExpandPasswordDispatchEntries(queue);
                foreach (SeparatePasswordDispatchEntry entry in expanded)
                {
                    Outlook.MailItem draft =
                        _owner.OutlookApplication.CreateItem(
                            Outlook.OlItemType.olMailItem)
                        as Outlook.MailItem;
                    if (draft == null)
                    {
                        throw new InvalidOperationException(
                            "A password draft could not be created.");
                    }
                    created.Add(draft);
                    _compose.PreparePasswordDraft(
                        draft,
                        entry,
                        composeKey);
                    WriteProperty(draft, AttemptProperty, attempt);
                    WriteProperty(draft, StateProperty, "staging");
                    WriteProperty(
                        draft,
                        SentFolderProperty,
                        sentFolderId);
                    WriteProperty(
                        draft,
                        SentStoreProperty,
                        sentStoreId);
                    WriteProperty(
                        draft,
                        PayloadProperty,
                        Protect(entry));
                    draft.Save();
                }

                WriteProperty(primary, AttemptProperty, attempt);
                if (!string.Equals(
                        ReadProperty(primary, AttemptProperty),
                        attempt,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Outlook did not retain the send-attempt marker.");
                }
                foreach (Outlook.MailItem draft in created)
                {
                    WriteProperty(draft, StateProperty, ArmedState);
                    draft.Save();
                    AddDraftReference(
                        attempt,
                        ReadEntryId(draft),
                        ReadParentStoreId(draft),
                        sentFolderId,
                        sentStoreId);
                }
                HookFolder(sentFolder);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Password-draft persistence failed; primary send cancelled (composeKey="
                    + (composeKey ?? string.Empty)
                    + ").",
                    ex);
                TryWriteProperty(primary, AttemptProperty, string.Empty);
                foreach (Outlook.MailItem draft in created)
                {
                    try
                    {
                        WriteProperty(draft, StateProperty, "cancelled");
                        draft.Delete();
                    }
                    catch
                    {
                        // A staging draft is ignored even if Outlook refuses
                        // its targeted cleanup.
                    }
                }
                return false;
            }
            finally
            {
                foreach (Outlook.MailItem draft in created)
                {
                    Release(draft, "prepared password draft");
                }
                Release(sentFolder, "exact Sent folder");
            }
        }

        private Outlook.MAPIFolder ResolveExactSentFolder(
            Outlook.MailItem primary)
        {
            Outlook.MAPIFolder folder = primary.SaveSentMessageFolder;
            if (folder != null)
            {
                return folder;
            }

            Outlook.Account account = null;
            Outlook.Store store = null;
            Outlook.NameSpace session = null;
            try
            {
                account = primary.SendUsingAccount;
                store = account != null ? account.DeliveryStore : null;
                folder = store != null
                    ? store.GetDefaultFolder(
                        Outlook.OlDefaultFolders.olFolderSentMail)
                    : null;
                if (folder == null)
                {
                    session = _owner.OutlookApplication.Session;
                    folder = session.GetDefaultFolder(
                        Outlook.OlDefaultFolders.olFolderSentMail);
                }
                if (folder == null)
                {
                    throw new InvalidOperationException(
                        "No Sent folder is available.");
                }
                primary.SaveSentMessageFolder = folder;
                return folder;
            }
            finally
            {
                Release(session, "Sent-folder Session");
                Release(store, "Sent-folder Store");
                Release(account, "Sent-folder Account");
            }
        }

        private void ScanDraftFolder(
            Outlook.MAPIFolder folder,
            Outlook.NameSpace session,
            HashSet<string> attempts)
        {
            Outlook.Items items = null;
            Outlook.Items filtered = null;
            try
            {
                if (!FolderDefinesProperty(folder, StateProperty))
                {
                    return;
                }
                items = folder != null ? folder.Items : null;
                filtered = items != null
                    ? items.Restrict(
                        BuildTextPropertyFilter(
                            StateProperty,
                            ArmedState))
                    : null;
                int count = filtered != null ? filtered.Count : 0;
                for (int i = 1; i <= count; i++)
                {
                    object item = null;
                    try
                    {
                        item = filtered[i];
                        Outlook.MailItem draft = item as Outlook.MailItem;
                        if (draft == null
                            || !string.Equals(
                                ReadProperty(draft, StateProperty),
                                ArmedState,
                                StringComparison.Ordinal))
                        {
                            continue;
                        }
                        string attempt =
                            ReadProperty(draft, AttemptProperty);
                        string sentFolderId =
                            ReadProperty(draft, SentFolderProperty);
                        string sentStoreId =
                            ReadProperty(draft, SentStoreProperty);
                        if (string.IsNullOrWhiteSpace(attempt)
                            || string.IsNullOrWhiteSpace(sentFolderId)
                            || string.IsNullOrWhiteSpace(sentStoreId))
                        {
                            continue;
                        }
                        AddDraftReference(
                            attempt,
                            ReadEntryId(draft),
                            ReadParentStoreId(draft),
                            sentFolderId,
                            sentStoreId);
                        attempts.Add(attempt);
                        HookFolder(session, sentFolderId, sentStoreId);
                    }
                    finally
                    {
                        Release(item, "recovery draft item");
                    }
                }
            }
            finally
            {
                Release(filtered, "filtered recovery Drafts items");
                Release(items, "recovery Drafts items");
            }
        }

        private static bool FolderDefinesProperty(
            Outlook.MAPIFolder folder,
            string propertyName)
        {
            Outlook.UserDefinedProperties properties = null;
            Outlook.UserDefinedProperty property = null;
            try
            {
                properties = folder != null
                    ? folder.UserDefinedProperties
                    : null;
                property = properties != null
                    ? properties.Find(propertyName)
                    : null;
                return property != null;
            }
            finally
            {
                Release(property, "UserDefinedProperty");
                Release(properties, "UserDefinedProperties");
            }
        }

        private void HookFolder(
            Outlook.NameSpace session,
            string entryId,
            string storeId)
        {
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = session.GetFolderFromID(entryId, storeId);
                HookFolder(folder);
            }
            finally
            {
                Release(folder, "resolved Sent folder");
            }
        }

        private void HookFolder(Outlook.MAPIFolder folder)
        {
            string key = FolderKey(folder.EntryID, folder.StoreID);
            if (_folderWatches.ContainsKey(key))
            {
                return;
            }
            Outlook.Items items = folder.Items;
            if (items == null)
            {
                throw new InvalidOperationException(
                    "Sent folder items are unavailable.");
            }
            items.ItemAdd += OnSentItemAdded;
            _folderWatches.Add(key, new FolderWatch(items));
        }

        private void OnSentItemAdded(object item)
        {
            Outlook.MailItem mail = item as Outlook.MailItem;
            if (_disposed || mail == null || !ReadSent(mail))
            {
                return;
            }
            string attempt = ReadProperty(mail, AttemptProperty);
            if (!string.IsNullOrWhiteSpace(attempt)
                && IsExpectedFolder(attempt, mail))
            {
                DispatchAttempt(attempt);
            }
        }

        private bool IsPrimarySent(
            Outlook.NameSpace session,
            string attempt)
        {
            List<DraftReference> references;
            if (!_drafts.TryGetValue(attempt, out references)
                || references.Count == 0)
            {
                return false;
            }
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;
            Outlook.Items filtered = null;
            try
            {
                DraftReference reference = references[0];
                folder = session.GetFolderFromID(
                    reference.SentFolderId,
                    reference.SentStoreId);
                items = folder != null ? folder.Items : null;
                filtered = items != null
                    ? items.Restrict(BuildAttemptFilter(attempt))
                    : null;
                int count = filtered != null ? filtered.Count : 0;
                for (int i = 1; i <= count; i++)
                {
                    object item = null;
                    try
                    {
                        item = filtered[i];
                        if (ReadSent(item as Outlook.MailItem))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        Release(item, "filtered Sent item");
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Exact Sent-folder recovery check failed; pending drafts retained.",
                    ex);
                return false;
            }
            finally
            {
                Release(filtered, "filtered Sent items");
                Release(items, "Sent items");
                Release(folder, "Sent folder");
            }
        }

        private bool IsExpectedFolder(
            string attempt,
            Outlook.MailItem sent)
        {
            List<DraftReference> references;
            if (!_drafts.TryGetValue(attempt, out references))
            {
                return false;
            }
            string entryId;
            string storeId;
            ReadParentFolderIds(sent, out entryId, out storeId);
            foreach (DraftReference reference in references)
            {
                if (string.Equals(
                        reference.SentFolderId,
                        entryId,
                        StringComparison.Ordinal)
                    && string.Equals(
                        reference.SentStoreId,
                        storeId,
                        StringComparison.Ordinal))
                {
                    return true;
                }
            }
            return false;
        }

        private void DispatchAttempt(string attempt)
        {
            List<DraftReference> references;
            if (!_dispatching.Add(attempt)
                || !_drafts.TryGetValue(attempt, out references))
            {
                return;
            }
            try
            {
                int recipients = 0;
                foreach (DraftReference reference
                    in new List<DraftReference>(references))
                {
                    Outlook.NameSpace session = null;
                    Outlook.MailItem draft = null;
                    try
                    {
                        session = _owner.OutlookApplication.Session;
                        draft = session.GetItemFromID(
                            reference.EntryId,
                            reference.StoreId)
                            as Outlook.MailItem;
                        if (draft == null
                            || !string.Equals(
                                ReadProperty(draft, StateProperty),
                                ArmedState,
                                StringComparison.Ordinal))
                        {
                            continue;
                        }
                        SeparatePasswordDispatchEntry entry =
                            Unprotect(ReadProperty(
                                draft,
                                PayloadProperty));
                        bool submitted;
                        bool manual;
                        int count;
                        bool sent = _compose.SendPasswordDraft(
                            draft,
                            entry,
                            attempt,
                            out submitted,
                            out manual,
                            out count);
                        if (sent)
                        {
                            recipients += count;
                        }
                        else if (submitted || manual)
                        {
                            TryWriteProperty(
                                draft,
                                StateProperty,
                                submitted ? "submitted" : "manual");
                            try
                            {
                                draft.Save();
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Pending password draft dispatch failed; draft retained.",
                            ex);
                    }
                    finally
                    {
                        Release(draft, "pending password draft");
                        Release(session, "dispatch Session");
                    }
                }
                _drafts.Remove(attempt);
                if (recipients > 0)
                {
                    _owner.ShowPasswordMailSuccessNotification(recipients);
                }
            }
            finally
            {
                _dispatching.Remove(attempt);
            }
        }

        private void AddDraftReference(
            string attempt,
            string entryId,
            string storeId,
            string sentFolderId,
            string sentStoreId)
        {
            if (string.IsNullOrWhiteSpace(entryId)
                || string.IsNullOrWhiteSpace(storeId))
            {
                throw new InvalidOperationException(
                    "The saved password draft has no stable identity.");
            }
            List<DraftReference> references;
            if (!_drafts.TryGetValue(attempt, out references))
            {
                references = new List<DraftReference>();
                _drafts.Add(attempt, references);
            }
            references.Add(new DraftReference
            {
                EntryId = entryId,
                StoreId = storeId,
                SentFolderId = sentFolderId,
                SentStoreId = sentStoreId
            });
        }

        private static string Protect(
            SeparatePasswordDispatchEntry entry)
        {
            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = 1024 * 1024
            };
            string json = serializer.Serialize(Payload.From(entry));
            byte[] protectedBytes = ProtectedData.Protect(
                Encoding.UTF8.GetBytes(json),
                Entropy,
                DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        private static SeparatePasswordDispatchEntry Unprotect(
            string value)
        {
            byte[] clear = ProtectedData.Unprotect(
                Convert.FromBase64String(value ?? string.Empty),
                Entropy,
                DataProtectionScope.CurrentUser);
            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = 1024 * 1024
            };
            Payload payload = serializer.Deserialize<Payload>(
                Encoding.UTF8.GetString(clear));
            if (payload == null)
            {
                throw new InvalidOperationException(
                    "Password draft payload is invalid.");
            }
            return payload.ToEntry();
        }

        private static void WriteProperty(
            Outlook.MailItem mail,
            string name,
            string value)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = mail.UserProperties;
                property = properties != null
                    ? properties.Find(name, true)
                    : null;
                if (property == null && properties != null)
                {
                    property = properties.Add(
                        name,
                        Outlook.OlUserPropertyType.olText,
                        true,
                        Type.Missing);
                }
                if (property == null)
                {
                    throw new InvalidOperationException(
                        "Outlook custom property could not be created.");
                }
                property.Value = value ?? string.Empty;
            }
            finally
            {
                Release(property, "UserProperty");
                Release(properties, "UserProperties");
            }
        }

        private static void TryWriteProperty(
            Outlook.MailItem mail,
            string name,
            string value)
        {
            try
            {
                if (mail != null)
                {
                    WriteProperty(mail, name, value);
                }
            }
            catch
            {
            }
        }

        private static string ReadProperty(
            Outlook.MailItem mail,
            string name)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = mail != null ? mail.UserProperties : null;
                property = properties != null
                    ? properties.Find(name, true)
                    : null;
                object value = property != null ? property.Value : null;
                return value != null ? value.ToString().Trim() : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                Release(property, "UserProperty");
                Release(properties, "UserProperties");
            }
        }

        private static string ReadEntryId(Outlook.MailItem mail)
        {
            return mail != null ? (mail.EntryID ?? string.Empty) : string.Empty;
        }

        private static string ReadParentStoreId(Outlook.MailItem mail)
        {
            string entryId;
            string storeId;
            ReadParentFolderIds(mail, out entryId, out storeId);
            return storeId;
        }

        private static void ReadParentFolderIds(
            Outlook.MailItem mail,
            out string entryId,
            out string storeId)
        {
            object parent = null;
            try
            {
                parent = mail != null ? mail.Parent : null;
                Outlook.MAPIFolder folder = parent as Outlook.MAPIFolder;
                entryId = folder != null
                    ? (folder.EntryID ?? string.Empty)
                    : string.Empty;
                storeId = folder != null
                    ? (folder.StoreID ?? string.Empty)
                    : string.Empty;
            }
            finally
            {
                Release(parent, "parent folder");
            }
        }

        private static bool ReadSent(Outlook.MailItem mail)
        {
            try
            {
                return mail != null && mail.Sent;
            }
            catch
            {
                return false;
            }
        }

        private static string BuildAttemptFilter(string attempt)
        {
            return BuildTextPropertyFilter(
                AttemptProperty,
                attempt);
        }

        private static string BuildTextPropertyFilter(
            string propertyName,
            string value)
        {
            const string uri =
                "http://schemas.microsoft.com/mapi/string/"
                + "{00020329-0000-0000-C000-000000000046}/";
            return "@SQL=\"" + uri
                   + (propertyName ?? string.Empty)
                   + "\" = '"
                   + (value ?? string.Empty).Replace("'", "''")
                   + "'";
        }

        private static string FolderKey(
            string entryId,
            string storeId)
        {
            return (storeId ?? string.Empty)
                   + "\n"
                   + (entryId ?? string.Empty);
        }

        private static void Release(object value, string label)
        {
            ComInteropScope.TryRelease(
                value,
                LogCategories.FileLink,
                "Failed to release " + label + ".");
        }

        public void Dispose()
        {
            _disposed = true;
            foreach (FolderWatch watch in _folderWatches.Values)
            {
                try
                {
                    watch.Items.ItemAdd -= OnSentItemAdded;
                }
                catch
                {
                }
                Release(watch.Items, "subscribed Sent items");
            }
            _folderWatches.Clear();
            _drafts.Clear();
        }

        private sealed class FolderWatch
        {
            internal FolderWatch(Outlook.Items items)
            {
                Items = items;
            }

            internal Outlook.Items Items { get; private set; }
        }

        private sealed class DraftReference
        {
            internal string EntryId { get; set; }
            internal string StoreId { get; set; }
            internal string SentFolderId { get; set; }
            internal string SentStoreId { get; set; }
        }

        private sealed class Payload
        {
            public string ShareLabel { get; set; }
            public string ShareUrl { get; set; }
            public string ShareId { get; set; }
            public string ShareToken { get; set; }
            public string RelativePath { get; set; }
            public DateTime? ExpireDate { get; set; }
            public int Permissions { get; set; }
            public string Password { get; set; }
            public string Html { get; set; }
            public string PlainText { get; set; }
            public string SecretsHtmlTemplate { get; set; }
            public string SecretsPlainTextTemplate { get; set; }
            public bool IsPlainText { get; set; }
            public int DeliveryMode { get; set; }
            public int SecretsExpireDays { get; set; }
            public string LanguageOverride { get; set; }
            public string To { get; set; }
            public string Cc { get; set; }
            public string Bcc { get; set; }
            public string SenderEmail { get; set; }
            public string SendUsingAccountSmtpAddress { get; set; }
            public string SentOnBehalfOfName { get; set; }
            public bool SignatureActive { get; set; }
            public string SignatureUserEmail { get; set; }
            public string SignatureHtml { get; set; }
            public string SignaturePlainText { get; set; }
            public ComposeLifecycleOrigin Origin { get; set; }

            internal static Payload From(
                SeparatePasswordDispatchEntry entry)
            {
                return new Payload
                {
                    ShareLabel = entry.ShareLabel,
                    ShareUrl = entry.ShareUrl,
                    ShareId = entry.ShareId,
                    ShareToken = entry.ShareToken,
                    RelativePath = entry.RelativePath,
                    ExpireDate = entry.ExpireDate,
                    Permissions = (int)entry.Permissions,
                    Password = entry.Password,
                    Html = entry.Html,
                    PlainText = entry.PlainText,
                    SecretsHtmlTemplate = entry.SecretsHtmlTemplate,
                    SecretsPlainTextTemplate =
                        entry.SecretsPlainTextTemplate,
                    IsPlainText = entry.IsPlainText,
                    DeliveryMode = (int)entry.DeliveryMode,
                    SecretsExpireDays = entry.SecretsExpireDays,
                    LanguageOverride = entry.LanguageOverride,
                    To = entry.To,
                    Cc = entry.Cc,
                    Bcc = entry.Bcc,
                    SenderEmail = entry.SenderEmail,
                    SendUsingAccountSmtpAddress =
                        entry.SendUsingAccountSmtpAddress,
                    SentOnBehalfOfName = entry.SentOnBehalfOfName,
                    SignatureActive = entry.SignatureActive,
                    SignatureUserEmail = entry.SignatureUserEmail,
                    SignatureHtml = entry.SignatureHtml,
                    SignaturePlainText = entry.SignaturePlainText,
                    Origin = entry.Origin != null
                        ? entry.Origin.Clone()
                        : null
                };
            }

            internal SeparatePasswordDispatchEntry ToEntry()
            {
                return new SeparatePasswordDispatchEntry
                {
                    ShareLabel = ShareLabel,
                    ShareUrl = ShareUrl,
                    ShareId = ShareId,
                    ShareToken = ShareToken,
                    RelativePath = RelativePath,
                    ExpireDate = ExpireDate,
                    Permissions = (FileLinkPermissionFlags)Permissions,
                    Password = Password,
                    Html = Html,
                    PlainText = PlainText,
                    SecretsHtmlTemplate = SecretsHtmlTemplate,
                    SecretsPlainTextTemplate =
                        SecretsPlainTextTemplate,
                    IsPlainText = IsPlainText,
                    DeliveryMode =
                        (SharePasswordDeliveryMode)DeliveryMode,
                    SecretsExpireDays = SecretsExpireDays,
                    LanguageOverride = LanguageOverride,
                    To = To,
                    Cc = Cc,
                    Bcc = Bcc,
                    SenderEmail = SenderEmail,
                    SendUsingAccountSmtpAddress =
                        SendUsingAccountSmtpAddress,
                    SentOnBehalfOfName = SentOnBehalfOfName,
                    SignatureActive = SignatureActive,
                    SignatureUserEmail = SignatureUserEmail,
                    SignatureHtml = SignatureHtml,
                    SignaturePlainText = SignaturePlainText,
                    Origin = Origin != null ? Origin.Clone() : null
                };
            }
        }
    }
}
