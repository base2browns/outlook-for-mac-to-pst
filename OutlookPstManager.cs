using Redemption;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Text.RegularExpressions;

namespace ConsoleApp1
{
    sealed class OutlookPstManager: IDisposable
    {
        private readonly RDOSession _session;
        private readonly RDOFolder _baseFolder;
        private readonly Dictionary<string, RDOFolder> _cachedFolders = new Dictionary<string, RDOFolder>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<int, Thread> _threads = new Dictionary<int, Thread>();

        public OutlookPstManager(Redemption.RDOSession session, Redemption.RDOFolder baseFolder)
        {
            this._session = session;
            this._baseFolder = baseFolder;
        }

        internal static OutlookPstManager Create(string storeId, string folderId)
        {
            var session = new Redemption.RDOSession();

            session.Logon();

            Redemption.RDOFolder folder;
            if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(folderId))
            {
                folder = session.PickFolder();
                if (folder == null) throw new Exception("Target folder not selected");
                Console.WriteLine($"Selected folder is {folder.EntryID} in store {folder.StoreID}");
            }
            else
            {
                folder = session.GetFolderFromID(folderId, storeId);
                if (folder == null) throw new ArgumentException($"Unknown target folder {folderId}", nameof(folderId));
            }

            return new OutlookPstManager(session, folder);
        }

        public string GetMessageId(FileInfo msgFile)
        {
            try
            {
                var email = MsgReader.Mime.Message.Load(msgFile);

                var msgId = email.Headers.MessageId;
                if (msgId != null && msgId.StartsWith("<") && msgId.EndsWith(">"))
                {
                    return msgId.Substring(1, msgId.Length - 2);
                }
                else
                {
                    return msgId;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to read email {msgFile.Name}: {ex.Message}");
                return null;
            }
        }

        public void AddEmailToFolder(string emailPath, int threadId, bool isFirstMsg, IReadOnlyList<string> folders)
        {
            var folder = _baseFolder;
            if (folders != null)
            {
                var path = string.Join("/", folders);
                if (!_cachedFolders.TryGetValue(path, out folder))
                {
                    folder = FindFolder(_baseFolder, folders, 0);
                    _cachedFolders.Add(path, folder);
                }
            }

            var msg = folder.Items.Add();

            if (threadId > 0)
            {
                if (!_threads.TryGetValue(threadId, out var thread))
                {
                    thread = new Thread();
                    _threads.Add(threadId, thread);
                }

                if (isFirstMsg)
                {
                    thread.FirstMessage = msg;

                    msg.CreateConversationIndex(msg);
                    msg.ConversationTopic = msg.Subject;

                    if (thread.PendingMessages != null)
                    {
                        foreach (var pendingMsg in thread.PendingMessages)
                        {
                            pendingMsg.CreateConversationIndex(msg);
                            pendingMsg.ConversationTopic = msg.Subject;
                            pendingMsg.Save();
                        }

                        thread.PendingMessages = null;
                    }
                }
                else if (thread.FirstMessage != null)
                {
                    msg.CreateConversationIndex(thread.FirstMessage);
                    msg.ConversationTopic = thread.FirstMessage.Subject;
                }
                else
                {
                    if (thread.PendingMessages == null) thread.PendingMessages = new List<RDOMail>();
                    thread.PendingMessages.Add(msg);
                }
            }

            msg.Sent = true;
            msg.Import(emailPath, 1024);
            msg.Save();
        }

        private RDOFolder FindFolder(RDOFolder folder, IReadOnlyList<string> folders, int index)
        {
            if (index >= folders.Count) return folder;

            var name = folders[index];
            
            foreach (RDOFolder child in folder.Folders)
            {
                if (string.Equals(child.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return FindFolder(child, folders, index + 1);
                }
            }

            var f = folder.Folders.Add(name);
            f.Save();

            return FindFolder(f, folders, index + 1);
        }

        public void Dispose()
        {
            if (_session.FastShutdownSupported)
            {
                _session.DoFastShutdown();
            }
            else
            {
                _session.Logoff();
            }
        }

        private sealed class Thread
        {
            public RDOMail FirstMessage { get; set; }

            public List<RDOMail> PendingMessages { get; set; }
        }
    }
}
