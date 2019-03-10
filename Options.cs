using System;
using System.IO;
using System.Windows.Forms;

namespace LnkUtils
{
	public static class Options
	{
		public static void Usage()
		{
			Log.Message("Usage: "+nameof(LnkUtils)+" (command) [options]"
				+"\n Commands:"
				+"\n"
				+"\n  Check (path to folder or .lnk file) [options]"
				+"\n  - tests link files for bad paths"
				+"\n"
				+"\n  Check Options:"
				+"\n  -r        recurse into sub-folders"
				+"\n  -v        print extra information about links"
				+"\n  -a        show all links not just bad ones"
				+"\n"
				+"\n  Create [options]"
				+"\n  - create a new link"
				+"\n"
				+"\n  Create Options:"
				+"\n  -t (path)          path to target"
				+"\n  -n (path)          name of link file (defaults to 'target file name'.lnk)"
				+"\n  -c (text)          comment"
				+"\n  -s (path)          start in path"
				// +"\n  -k (key)           shortcut key"
				// +"\n  -r (run)           starting window state"
				+"\n  -i (path) [index]  path to .ico file or .dll/.exe. index defaults to 0"
				// +"\n  -a                 enable run as administrator"
				+"\n  -f                 force create even if target doesn't exist"
			);
		}

		public static bool Parse(string[] args)
		{
			// System.Windows.Forms.Keys
			int len = args.Length;
			for(int a=0; a<len; a++)
			{
				string curr = args[a];
				if (Action == Command.None) {
					if (!Enum.TryParse<Command>(curr ?? "",true,out Action)) {
						Log.Error("Invalid action "+curr);
						return false;
					}
				}
				else if (Action == Command.Check) {
					if (curr == "-r") {
						Recurse = true;
					}
					else if (curr == "-v") {
						Verbose = true;
					}
					else if (curr == "-a") {
						ShowAll = true;
					}
					else {
						Target = curr;
					}
				}
				else if (Action == Command.Create) {
					if (curr == "-t" && ++a < len) {
						Target = Path.GetFullPath(args[a]);
					}
					else if (curr == "-n" && ++a < len) {
						LnkFileName = args[a];
					}
					else if (curr == "-c" && ++a < len) {
						Comment = args[a];
					}
					else if (curr == "-s" && ++a < len) {
						StartIn = args[a];
					}
					else if (curr == "k" && ++a < len) {
						//TODO
					}
					else if (curr == "-r" && ++a < len) {
						//TODO
					}
					else if (curr == "-i" && ++a < len) {
						string snum = args[a+1];
						IconPath = args[a];
						if (int.TryParse(snum,out int num)) {
							IconIndex = num; ++a;
						}
					}
					else if (curr == "-a") {
						EnableAdmin = true;
						//TODO not used yet
					}
					else if (curr == "-f") {
						Force = true;
					}
				}
			}
			return true;
		}

		public enum Command
		{
			None = 0,
			Check = 1,
			Create = 2
		}

		public static Command Action = Command.None;
		public static bool Recurse = false;
		public static bool Verbose = false;
		public static bool ShowAll = false;
		public static string Target = null;
		public static string Comment = null;
		public static string StartIn = null;
		public static string IconPath = null;
		public static int IconIndex = 0;
		public static bool EnableAdmin = false;
		public static string LnkFileName = null;
		public static bool Force = false;

	}
}