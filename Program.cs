﻿using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices.ComTypes;

namespace LnkUtils
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length < 1) {
				Options.Usage();
				return;
			}
			if (!Options.Parse(args)) {
				return;
			}

			try {
				DoActions();
			} catch(Exception e) {
				#if DEBUG
				Log.Error(e.ToString());
				#else
				log.Error(e.Message);
				#endif
			}
		}

		static void DoActions()
		{

			if (Options.Action == Options.Command.Check) {
				ActionCheck();
			}
			else if (Options.Action == Options.Command.Create) {
				ActionCreate();
			}
		}

		static void ActionCheck()
		{
			bool isFolder = false;
			string target = Options.Target;

			if (Directory.Exists(target)) {
				isFolder = true;
			}
			else if (Path.GetExtension(target) != ".lnk" || !File.Exists(target)) {
				Log.Error("invalid shortcut file: "+target);
				return;
			}

			int count = 0;
			if (isFolder) {
				SearchOption so = Options.Recurse
					? SearchOption.AllDirectories
					: SearchOption.TopDirectoryOnly
				;
				var iter = Directory.GetFiles(target,"*.lnk",so);
				foreach(string linkFile in iter) {
					Check(linkFile);
					count++;
				}
			}
			else {
				Check(target);
				count++;
			}
			Log.Message("Checked "+count+" shortcuts");
		}

		static void Check(string lnkFile)
		{
			IShellLinkW link = (IShellLinkW)new ShellLink();
			((IPersistFile)link).Load(lnkFile,(int)STGM_FLAGS.STGM_READ);
			string path = GetPath(link, out WIN32_FIND_DATA pathData);
			string wdir = GetWorkingDirectory(link);

			bool bad = false;
			if (!File.Exists(path)) {
				bad = true;
				Log.Message("Bad Target  : "+lnkFile);
			}

			if (!File.Exists(path)) {
				bad = true;
				Log.Message("Bad Start In: "+lnkFile);
			}

			if (!bad && Options.ShowAll) {
				Log.Message("Valid       : "+lnkFile);
			}

			if (Options.Verbose) {
				PrintExtraData(path,wdir,pathData,bad);
			}
		}

		static string GetPath(IShellLinkW link, out WIN32_FIND_DATA data)
		{
			var sb = new StringBuilder(16384);
			link.GetPath(sb,sb.Capacity,out data, SLGP_FLAGS.SLGP_NONE);
			return sb.ToString();
		}

		static string GetWorkingDirectory(IShellLinkW link)
		{
			var sb = new StringBuilder(16384);
			link.GetWorkingDirectory(sb,sb.Capacity);
			return sb.ToString();
		}

		static void PrintExtraData(string path,string wdir,WIN32_FIND_DATA data, bool isBad)
		{
			long size = ((long)data.nFileSizeHigh << 32) + data.nFileSizeLow;
			Log.Message(""
				+   "Target      : "+(isBad ? "" : path)
				+ "\nStart In    : "+(isBad ? "" : wdir)
				+ "\nFileName    : "+(isBad ? "" : data.cFileName)
				+ "\n8.3 FileName: "+(isBad ? "" : data.cAlternateFileName)
				+ "\nSize        : "+(isBad ? "" : size.ToString("N0"))
				+ "\nCreated     : "+(isBad ? "" : ConvertFileTime(data.ftCreationTime))
				+ "\nModified    : "+(isBad ? "" : ConvertFileTime(data.ftLastWriteTime))
				+ "\nAccessed    : "+(isBad ? "" : ConvertFileTime(data.ftLastAccessTime))
				+ "\nAttributes  : "+(isBad ? "" : data.dwFileAttributes.ToString())
				+ "\n"
			);
		}

		static string ConvertFileTime(System.Runtime.InteropServices.ComTypes.FILETIME ft)
		{
			long time = ((long)ft.dwHighDateTime << 32) + ft.dwLowDateTime;
			var dto = DateTimeOffset.FromFileTime(time).ToLocalTime();
			return dto.ToString();
		}

		static bool EndsWithIC(string subj, string test)
		{
			if (subj == null || test == null) { return false; }
			return subj.EndsWith(test,StringComparison.OrdinalIgnoreCase);
		}

		static void ActionCreate()
		{
			if (String.IsNullOrWhiteSpace(Options.Target)) {
				Log.Error("invalid path to target");
				return;
			}
			bool isDirectory = false;
			Log.Debug("force = "+Options.Force);
			if (!Options.Force && !Exists(Options.Target,out isDirectory)) {
				Log.Error("path to target doesn't exist");
				return;
			}
			if (String.IsNullOrWhiteSpace(Options.LnkFileName)) {
				Options.LnkFileName = Path.GetFileNameWithoutExtension(Options.Target);
			}
			if (!EndsWithIC(Options.LnkFileName,".lnk")) {
				Options.LnkFileName += ".lnk";
			}

			IShellLinkW link = (IShellLinkW)new ShellLink();
			if (!String.IsNullOrWhiteSpace(Options.Comment)) {
				link.SetDescription(Options.Comment);
			}
			if (!String.IsNullOrWhiteSpace(Options.IconPath)) {
				link.SetIconLocation(Options.IconPath,Options.IconIndex);
			}
			if (!String.IsNullOrWhiteSpace(Options.StartIn)) {
				link.SetWorkingDirectory(Options.StartIn);
			}
			link.SetPath(Options.Target);

			Log.Debug("saving "+Options.LnkFileName);
			((IPersistFile)link).Save(Options.LnkFileName,false);
		}

		static bool Exists(string path, out bool isDirectory)
		{
			isDirectory = false;
			if (String.IsNullOrEmpty(path)) {
				return false;
			}
			else if (File.Exists(path)) {
				Log.Debug("File exists "+path);
				return true;
			}
			else if (Directory.Exists(path)) {
				Log.Debug("Directory exists "+path);
				isDirectory = true;
				return true;
			}
			else if (
				path.EndsWith(Path.DirectorySeparatorChar.ToString())
				|| path.EndsWith(Path.AltDirectorySeparatorChar.ToString())
			) {
				isDirectory = true;
			}
			Log.Debug("does not exist "+path);
			return false;
		}
	}
}
