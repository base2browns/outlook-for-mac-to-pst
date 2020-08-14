using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace ConsoleApp1
{
    class OutlookSqliteQuery
    {
        private Dictionary<int, Folder> _folders = new Dictionary<int, Folder>();
        private Dictionary<string, Mail> _messages = new Dictionary<string, Mail>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<int, Thread> _threads = new Dictionary<int, Thread>();
        
        public int MessageCount => _messages.Count;

        public OutlookSqliteQuery(string sqlitePath)
        {
            Console.WriteLine("Loading data ...");
            using (var conn = new SQLiteConnection($"Data Source={sqlitePath}"))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    Console.WriteLine("Loading folders ...");
                    cmd.CommandText = "select Record_RecordId, Folder_ParentID, Folder_Name from Folders where Folder_Name is not null";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            _folders.Add(reader.GetInt32(0), new Folder
                            {
                                ParentID = reader.GetInt32(1),
                                Name = reader.GetString(2)
                            });
                        }
                    }

                    Console.WriteLine("Loading emails ...");
                    cmd.CommandText = "select Message_MessageID, Record_RecordID, Record_FolderID, Threads_ThreadID from Mail where Message_MessageID is not null";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var msgId = reader.GetString(0);
                            if (msgId.StartsWith("<") && msgId.EndsWith(">"))
                            {
                                msgId = msgId.Substring(1, msgId.Length - 2);
                            }

                            var mail = new Mail
                            {
                                RecordID = reader.GetInt32(1),
                                FolderID = reader.GetInt32(2),
                                ThreadID = reader.GetInt32(3)
                            };

                            if (_messages.TryAdd(msgId, mail))
                            {
                                if (_threads.TryGetValue(mail.ThreadID, out var thread))
                                {
                                    thread.Count++;
                                    if (thread.FirstRecordID > mail.RecordID)
                                    {
                                        thread.FirstRecordID = mail.RecordID;
                                    }
                                }
                                else
                                {
                                    _threads.Add(mail.ThreadID, new Thread { Count = 1, FirstRecordID = mail.RecordID });
                                }
                            }
                        }
                    }

                    Console.WriteLine("Loading data complete");
                }
            }
        }       

        internal static OutlookSqliteQuery Create(string profilePath)
        {
            var sqlitePath = Path.Combine(profilePath, "Outlook.sqlite");
            if (!File.Exists(sqlitePath)) throw new ArgumentException("Invalid outlook profile path", nameof(profilePath));
            return new OutlookSqliteQuery(sqlitePath);
        }

        public bool TryFindMessage(string msgId, out int threadId, out bool isFirstMsg, out IReadOnlyList<string> folderPath)
        {
            folderPath = null;
            threadId = -1;
            isFirstMsg = false;

            if (string.IsNullOrEmpty(msgId) || !_messages.TryGetValue(msgId, out var mail))
            {
                return false;
            }

            mail.Seen = true;

            if (_threads.TryGetValue(mail.ThreadID, out var thread) && thread.Count > 1)
            {
                threadId = mail.ThreadID;
                isFirstMsg = thread.FirstRecordID == mail.RecordID;
            }

            if (_folders.TryGetValue(mail.FolderID, out var folder))
            {
                if (folder.Path == null)
                {
                    var path = new List<string>();

                    var folderId = mail.FolderID;
                    while (_folders.TryGetValue(folderId, out var f))
                    {
                        path.Add(f.Name);
                        folderId = f.ParentID;
                    }
                    
                    path.Reverse();

                    folder.Path = path;
                }

                folderPath = folder.Path;
            }

            return true;
        }

        public sealed class Folder
        {
            public int ParentID { get; set; }
            public string Name { get; set; }

            public IReadOnlyList<string> Path { get; set; }
        }

        public sealed class Mail
        {
            public int RecordID { get; set; }
            public int FolderID { get; set; }
            public int ThreadID { get; set; }
            public bool Seen { get; set; }
        }

        public sealed class Thread
        {
            public int FirstRecordID { get; set; }
            public int Count { get; set; }
        }
    }
}
