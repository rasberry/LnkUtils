using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.IO;

// https://stackoverflow.com/questions/15418004/how-do-i-go-about-instantiating-a-com-object-in-c-sharp-by-clsid
// https://social.msdn.microsoft.com/Forums/vstudio/en-US/8d2dc042-01c6-4870-95b2-1ff0a72c398a/simulating-explorer-send-to-functionality-using-com-interop-in-cnet-35

namespace LnkUtils
{
	[ComImport]
	[Guid("00021401-0000-0000-C000-000000000046")]
	public class ShellLink
	{
	}

	[ComImport]
	[Guid("0AFACED1-E828-11D1-9187-B532F1E9575D")]
	public class ShellLinkFolder
	{
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("000214F9-0000-0000-C000-000000000046")]
	public interface IShellLink
	{
		void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, int fFlags);
		void GetIDList(out IntPtr ppidl);
		void SetIDList(IntPtr pidl);
		void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
		void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
		void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
		void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
		void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
		void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
		void GetHotkey(out short pwHotkey);
		void SetHotkey(short wHotkey);
		void GetShowCmd(out int piShowCmd);
		void SetShowCmd(int iShowCmd);
		void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
		void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
		void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
		[return: MarshalAs(UnmanagedType.U4)] uint Resolve(IntPtr hwnd, int fFlags);
		void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
	}

	/// <summary>The IShellLink interface allows Shell links to be created, modified, and resolved</summary>
	[ComImport()]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("000214F9-0000-0000-C000-000000000046")]
	interface IShellLinkW
	{
		/// <summary>Retrieves the path and file name of a Shell link object</summary>
		void GetPath([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out WIN32_FIND_DATA pfd, SLGP_FLAGS fFlags);
		/// <summary>Retrieves the list of item identifiers for a Shell link object</summary>
		void GetIDList(out IntPtr ppidl);
		/// <summary>Sets the pointer to an item identifier list (PIDL) for a Shell link object.</summary>
		void SetIDList(IntPtr pidl);
		/// <summary>Retrieves the description string for a Shell link object</summary>
		void GetDescription([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
		/// <summary>Sets the description for a Shell link object. The description can be any application-defined string</summary>
		void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
		/// <summary>Retrieves the name of the working directory for a Shell link object</summary>
		void GetWorkingDirectory([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
		/// <summary>Sets the name of the working directory for a Shell link object</summary>
		void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
		/// <summary>Retrieves the command-line arguments associated with a Shell link object</summary>
		void GetArguments([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
		/// <summary>Sets the command-line arguments for a Shell link object</summary>
		void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
		/// <summary>Retrieves the hot key for a Shell link object</summary>
		void GetHotkey(out short pwHotkey);
		/// <summary>Sets a hot key for a Shell link object</summary>
		void SetHotkey(short wHotkey);
		/// <summary>Retrieves the show command for a Shell link object</summary>
		void GetShowCmd(out int piShowCmd);
		/// <summary>Sets the show command for a Shell link object. The show command sets the initial show state of the window.</summary>
		void SetShowCmd(int iShowCmd);
		/// <summary>Retrieves the location (path and index) of the icon for a Shell link object</summary>
		void GetIconLocation([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
		/// <summary>Sets the location (path and index) of the icon for a Shell link object</summary>
		void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
		/// <summary>Sets the relative path to the Shell link object</summary>
		void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
		/// <summary>Attempts to find the target of a Shell link, even if it has been moved or renamed</summary>
		[return: MarshalAs(UnmanagedType.U4)] uint Resolve(IntPtr hwnd, SLR_FLAGS fFlags);
		/// <summary>Sets the path and file name of a Shell link object</summary>
		void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
	}

	// https://docs.microsoft.com/en-us/windows/desktop/Stg/stgm-constants
	[Flags]
	public enum STGM_FLAGS
	{
		STGM_READ = 0x0,
		STGM_WRITE = 0x1,
		STGM_READWRITE = 0x2,
		STGM_SHARE_DENY_NONE = 0x40,
		STGM_SHARE_DENY_READ = 0x30,
		STGM_SHARE_DENY_WRITE = 0x20,
		STGM_SHARE_EXCLUSIVE = 0x10,
		STGM_PRIORITY = 0x40000,
		STGM_CREATE = 0x1000,
		STGM_CONVERT = 0x20000,
		STGM_FAILIFTHERE = 0x0,
		STGM_DIRECT = 0x0,
		STGM_TRANSACTED = 0x10000,
		STGM_NOSCRATCH = 0x100000,
		STGM_NOSNAPSHOT = 0x200000,
		STGM_SIMPLE = 0x8000000,
		STGM_DIRECT_SWMR = 0x400000,
		STGM_DELETEONRELEASE = 0x4000000
	}

	[Flags]
	public enum SLR_FLAGS
	{
		SLR_NONE = 0x0,
		SLR_NO_UI = 0x1,
		SLR_ANY_MATCH = 0x2,
		SLR_NOUPDATE = 0x8,
		SLR_NOSEARCH = 0x10,
		SLR_NOTRACK = 0x20,
		SLR_NOLINKINFO = 0x40,
		SLR_INVOKE_MSI = 0x80
	}

	public enum HRESULT : uint
	{
		S_FALSE = 0x0001,
		S_OK = 0x0000,
		E_INVALIDARG = 0x80070057,
		E_OUTOFMEMORY = 0x8007000E
	}

	// The CharSet must match the CharSet of the corresponding PInvoke signature
	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
	struct WIN32_FIND_DATA
	{
		public FileAttributes dwFileAttributes;
		public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
		public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
		public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
		public uint nFileSizeHigh;
		public uint nFileSizeLow;
		public ReparseTagType dwReserved0;
		public uint dwReserved1;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst=260)]
		public string cFileName;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst=14)]
		public string cAlternateFileName;
	}

	/// <summary>IShellLink.GetPath fFlags: Flags that specify the type of path information to retrieve</summary>
	[Flags]
	enum SLGP_FLAGS {
		SLGP_NONE = 0x0,
		/// <summary>Retrieves the standard short (8.3 format) file name</summary>
		SLGP_SHORTPATH = 0x1,
		/// <summary>Retrieves the Universal Naming Convention (UNC) path name of the file</summary>
		SLGP_UNCPRIORITY = 0x2,
		/// <summary>Retrieves the raw path name. A raw path is something that might not exist and may include environment variables that need to be expanded</summary>
		SLGP_RAWPATH = 0x4,
		SLGP_RELATIVEPRIORITY = 0x8
	}

	/// <summary>
	/// File attributes are metadata values stored by the file system on disk and are used by the system and are available to developers via various file I/O APIs.
	/// </summary>
	[Flags]
	enum FileAttributes : uint
	{
		None = 0x00000000,

		/// <summary>
		/// A file that is read-only. Applications can read the file, but cannot write to it or delete it. This attribute is not honored on directories. For more information, see "You cannot view or change the Read-only or the System attributes of folders in Windows Server 2003, in Windows XP, or in Windows Vista".
		/// </summary>
		Readonly = 0x00000001,

		/// <summary>
		/// The file or directory is hidden. It is not included in an ordinary directory listing.
		/// </summary>
		Hidden = 0x00000002,

		/// <summary>
		/// A file or directory that the operating system uses a part of, or uses exclusively.
		/// </summary>
		System = 0x00000004,

		/// <summary>
		/// The handle that identifies a directory.
		/// </summary>
		Directory = 0x00000010,

		/// <summary>
		/// A file or directory that is an archive file or directory. Applications typically use this attribute to mark files for backup or removal.
		/// </summary>
		Archive = 0x00000020,

		/// <summary>
		/// This value is reserved for system use.
		/// </summary>
		Device = 0x00000040,

		/// <summary>
		/// A file that does not have other attributes set. This attribute is valid only when used alone.
		/// </summary>
		Normal = 0x00000080,

		/// <summary>
		/// A file that is being used for temporary storage. File systems avoid writing data back to mass storage if sufficient cache memory is available, because typically, an application deletes a temporary file after the handle is closed. In that scenario, the system can entirely avoid writing the data. Otherwise, the data is written after the handle is closed.
		/// </summary>
		Temporary = 0x00000100,

		/// <summary>
		/// A file that is a sparse file.
		/// </summary>
		SparseFile = 0x00000200,

		/// <summary>
		/// A file or directory that has an associated reparse point, or a file that is a symbolic link.
		/// </summary>
		ReparsePoint = 0x00000400,

		/// <summary>
		/// A file or directory that is compressed. For a file, all of the data in the file is compressed. For a directory, compression is the default for newly created files and subdirectories.
		/// </summary>
		Compressed = 0x00000800,

		/// <summary>
		/// The data of a file is not available immediately. This attribute indicates that the file data is physically moved to offline storage. This attribute is used by Remote Storage, which is the hierarchical storage management software. Applications should not arbitrarily change this attribute.
		/// </summary>
		Offline = 0x00001000,

		/// <summary>
		/// The file or directory is not to be indexed by the content indexing service.
		/// </summary>
		NotContentIndexed = 0x00002000,

		/// <summary>
		/// A file or directory that is encrypted. For a file, all data streams in the file are encrypted. For a directory, encryption is the default for newly created files and subdirectories.
		/// </summary>
		Encrypted = 0x00004000,

		/// <summary>
		/// This value is reserved for system use.
		/// </summary>
		Virtual = 0x00010000,

		/// <summary>
		/// When this attribute is set, it means that the file or directory has no physical representation on the local system; the item is virtual.
		/// </summary>
		RecallOnOpen = 0x00040000,

		/// <summary>
		/// When this attribute is set, it means that the file or directory is not fully present locally.
		/// </summary>
		RecallOnDataAccess = 0x00400000,
	}

	enum ReparseTagType : uint
	{
		IO_REPARSE_TAG_CSV = 0x80000009,
		IO_REPARSE_TAG_DEDUP = 0x80000013,
		IO_REPARSE_TAG_DFS = 0x8000000A,
		IO_REPARSE_TAG_DFSR = 0x80000012,
		IO_REPARSE_TAG_HSM = 0xC0000004,
		IO_REPARSE_TAG_HSM2 = 0x80000006,
		IO_REPARSE_TAG_MOUNT_POINT = 0xA0000003,
		IO_REPARSE_TAG_NFS = 0x80000014,
		IO_REPARSE_TAG_SIS = 0x80000007,
		IO_REPARSE_TAG_SYMLINK = 0xA000000C,
		IO_REPARSE_TAG_WIM = 0x80000008
	}
}