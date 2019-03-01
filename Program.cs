using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices.ComTypes;

namespace LnkUtils
{
	class Program
	{
		static void Main(string[] args)
		{
			string target = args.Length > 0 ? args[0] : null;
			bool isFolder = false;

			if (String.IsNullOrWhiteSpace(target)) {
				target = Environment.CurrentDirectory;
				isFolder = true;
			}
			else {
				if (Directory.Exists(target)) {
					isFolder = true;
				}
				else if (Path.GetExtension(target) != ".lnk") {
					Console.Error.Write("invalid shortcut file: "+target);
					return;
				}
			}

			int count = 0;
			if (isFolder) {
				var iter = Directory.GetFiles(target,"*.lnk");
				foreach(string linkFile in iter) {
					Check(linkFile);
					count++;
				}
			}
			else {
				Check(target);
				count++;
			}
			Console.WriteLine("Checked "+count+" shortcuts");
		}

		static void Check(string lnkFile)
		{
			if (!CheckShortcut(lnkFile)) {
				Console.WriteLine("Bad Link "+lnkFile);
			}
		}

		static bool CheckShortcut(string lnkPath)
		{
			if (!File.Exists(lnkPath)) {
				return false;
			}

			IShellLinkW link = (IShellLinkW)new ShellLink();
			((IPersistFile)link).Load(lnkPath,(int)STGM_FLAGS.STGM_READ);
			string path = GetPath(link);
			//Console.WriteLine("path = "+path);
			//var result = link.Resolve(IntPtr.Zero,SLR_FLAGS.SLR_NO_UI);
			//Console.WriteLine("Result = "+result);
			return File.Exists(path);
			// return result == 0;
		}

		static string GetPath(IShellLinkW link)
		{
			var sb = new StringBuilder(16384);
			link.GetPath(sb,sb.Capacity,out WIN32_FIND_DATA data, SLGP_FLAGS.SLGP_NONE);
			return sb.ToString();
		}

		static string GetWorkingDirectory(IShellLinkW link)
		{
			var sb = new StringBuilder(16384);
			link.GetWorkingDirectory(sb,sb.Capacity);
			return sb.ToString();
		}
	}
}


//WshShell wsh = new WshShell();
//IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(
//    Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\shorcut.lnk") as IWshRuntimeLibrary.IWshShortcut;
//shortcut.Arguments = "";
//shortcut.TargetPath = "c:\\app\\myftp.exe";
//// not sure about what this is for
//shortcut.WindowStyle = 1;
//shortcut.Description = "my shortcut description";
//shortcut.WorkingDirectory = "c:\\app";
//shortcut.IconLocation = "specify icon location";
//shortcut.Save();

//using Shell32;
//using IWshRuntimeLibrary;
//
//var wsh = new IWshShell_Class();
//IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(
//    Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\shorcut2.lnk") as IWshRuntimeLibrary.IWshShortcut;
//shortcut.TargetPath = @"C:\Users\Zimin\Desktop\test folder";
//shortcut.Save();
