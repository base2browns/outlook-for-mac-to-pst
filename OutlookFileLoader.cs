using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    class OutlookFileLoader
    {
        private readonly DirectoryInfo _messageDir;

        private OutlookFileLoader(DirectoryInfo messageDir)
        {
            _messageDir = messageDir;
        }

        internal static OutlookFileLoader Create(string profilePath)
        {
            var profileDir = new DirectoryInfo(profilePath);
            if (!profileDir.Exists) throw new ArgumentException("Profile path is not a directory", nameof(profilePath));

            var messageDir = profileDir.GetDirectories("Message Sources").FirstOrDefault();
            if (messageDir == null) throw new ArgumentException("Invalid profile path directory", nameof(profilePath));

            return new OutlookFileLoader(messageDir);
        }

        public IEnumerable<FileInfo> EnumerateMessages()
        {
            return _messageDir.EnumerateFiles("*.olk15MsgSource", new EnumerationOptions { RecurseSubdirectories = true });
        }
    }
}
