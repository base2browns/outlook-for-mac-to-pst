using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;

namespace ImportExport
{
    #region ManagedIStream object

    /// <summary>
    /// ManagedIStream wraps a .Net Stream object as a COM IStream interface
    /// </summary>
    /// 
    [Guid("0000000c-0000-0000-C000-000000000046"),
     ClassInterface(ClassInterfaceType.None)]
    public class ManagedIStream : IStream, IDisposable
    {
        public ManagedIStream(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream parameter cannot be null");
            _stream = stream;
        }

        [SecurityCritical]
        void IStream.Read(Byte[] buffer, Int32 bufferSize, IntPtr bytesReadPtr)
        {
            Int32 bytesRead = _stream.Read(buffer, 0, (int)bufferSize);
            if (bytesReadPtr != IntPtr.Zero)
            {
                Marshal.WriteInt32(bytesReadPtr, bytesRead);
            }
        }

        [SecurityCritical]
        void IStream.Seek(Int64 offset, Int32 origin, IntPtr newPositionPtr)
        {
            SeekOrigin seekOrigin;

            switch (origin)
            {
                case StreamConsts.STREAM_SEEK_SET:
                    seekOrigin = SeekOrigin.Begin;
                    break;
                case StreamConsts.STREAM_SEEK_CUR:
                    seekOrigin = SeekOrigin.Current;
                    break;
                case StreamConsts.STREAM_SEEK_END:
                    seekOrigin = SeekOrigin.End;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("origin parameter can only be STREAM_SEEK_SET / STREAM_SEEK_CUR / STREAM_SEEK_END");
            }
            long position = _stream.Seek(offset, seekOrigin);

            if (newPositionPtr != IntPtr.Zero)
            {
                Marshal.WriteInt64(newPositionPtr, position);
            }
        }

        void IStream.SetSize(Int64 libNewSize)
        {
            _stream.SetLength(libNewSize);
        }

        void IStream.Stat(out System.Runtime.InteropServices.ComTypes.STATSTG streamStats, int grfStatFlag)
        {
            streamStats = new System.Runtime.InteropServices.ComTypes.STATSTG();
            streamStats.type = (int)STGTY.STGTY_STREAM;
            streamStats.cbSize = _stream.Length;

            streamStats.grfMode = 0; // default value 
            if (_stream.CanRead && _stream.CanWrite)
            {
                streamStats.grfMode |= (int)STGM.READWRITE;
            }
            else if (_stream.CanRead)
            {
                streamStats.grfMode |= (int)STGM.READ;
            }
            else if (_stream.CanWrite)
            {
                streamStats.grfMode |= (int)STGM.WRITE;
            }
            else
            {
                // A stream that is neither readable nor writable is a closed stream.
                throw new IOException("Cannot access a closed stream");
            }
        }

        [SecurityCritical]
        void IStream.Write(Byte[] buffer, Int32 bufferSize, IntPtr bytesWrittenPtr)
        {
            _stream.Write(buffer, 0, bufferSize);
            if (bytesWrittenPtr != IntPtr.Zero)
            {
                // If fewer than bufferSize bytes had been written, an exception would
                // have been thrown, so it can be assumed we wrote bufferSize bytes.
                Marshal.WriteInt32(bytesWrittenPtr, bufferSize);
            }
        }

        public void Dispose()
        {
            _stream.Dispose();
            _stream = null;
        }

        #region Unimplemented methods

        void IStream.Clone(out IStream streamCopy)
        {
            streamCopy = null;
            throw new NotSupportedException();
        }

        void IStream.CopyTo(IStream targetStream, Int64 bufferSize, IntPtr buffer, IntPtr bytesWrittenPtr)
        {
            throw new NotSupportedException();
        }

        void IStream.Commit(Int32 flags)
        {
            //do nothing
        }

        void IStream.LockRegion(Int64 offset, Int64 byteCount, Int32 lockType)
        {
            throw new NotSupportedException();
        }

        void IStream.Revert()
        {
            throw new NotSupportedException();
        }

        void IStream.UnlockRegion(Int64 offset, Int64 byteCount, Int32 lockType)
        {
            throw new NotSupportedException();
        }


        #endregion Unimplemented methods

        #region Fields
        protected Stream _stream;
        #endregion Fields
    }

    #endregion ManagedIStream object

    #region FileIStream object

    /// <summary>
    /// FileIStream wraps a file as an IStream object
    /// </summary>
    public class FileIStream : ManagedIStream
    {
        public FileIStream(string FileName, bool CreateNew, bool Writable)
            : base(File.Open(FileName, CreateNew ? FileMode.Create : FileMode.Open, Writable ? FileAccess.ReadWrite : FileAccess.Read))
        {
            _stream.Position = 0;
        }
    }
    #endregion FileIStream object

    #region constants
    public sealed class StreamConsts
    {
        public const int LOCK_WRITE = 0x1;
        public const int LOCK_EXCLUSIVE = 0x2;
        public const int LOCK_ONLYONCE = 0x4;
        public const int STATFLAG_DEFAULT = 0x0;
        public const int STATFLAG_NONAME = 0x1;
        public const int STATFLAG_NOOPEN = 0x2;
        public const int STGC_DEFAULT = 0x0;
        public const int STGC_OVERWRITE = 0x1;
        public const int STGC_ONLYIFCURRENT = 0x2;
        public const int STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 0x4;
        public const int STREAM_SEEK_SET = 0x0;
        public const int STREAM_SEEK_CUR = 0x1;
        public const int STREAM_SEEK_END = 0x2;
    }

    [Flags]
    public enum STGM : int
    {
        DIRECT = 0x00000000,
        TRANSACTED = 0x00010000,
        SIMPLE = 0x08000000,
        READ = 0x00000000,
        WRITE = 0x00000001,
        READWRITE = 0x00000002,
        SHARE_DENY_NONE = 0x00000040,
        SHARE_DENY_READ = 0x00000030,
        SHARE_DENY_WRITE = 0x00000020,
        SHARE_EXCLUSIVE = 0x00000010,
        PRIORITY = 0x00040000,
        DELETEONRELEASE = 0x04000000,
        NOSCRATCH = 0x00100000,
        CREATE = 0x00001000,
        CONVERT = 0x00020000,
        FAILIFTHERE = 0x00000000,
        NOSNAPSHOT = 0x00200000,
        DIRECT_SWMR = 0x00400000,
    }

    public enum STATFLAG : uint
    {
        STATFLAG_DEFAULT = 0,
        STATFLAG_NONAME = 1,
        STATFLAG_NOOPEN = 2
    }

    public enum STGTY : int
    {
        STGTY_STORAGE = 1,
        STGTY_STREAM = 2,
        STGTY_LOCKBYTES = 3,
        STGTY_PROPERTY = 4
    }

    #endregion constants

}