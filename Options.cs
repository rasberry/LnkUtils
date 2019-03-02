using System;

namespace LnkUtils
{
	public static class Options
	{
		public static void Usage()
		{
			Log.Message("Usage: "+nameof(LnkUtils)+" [options] (path to folder or .lnk file)"
				+"\n Options:"
				+"\n  -r        recurse into sub-folders"
				+"\n  -v        print extra information about links"
				+"\n  -a        show all links not just bad ones"
			);
		}

		public static bool Parse(string[] args)
		{
			for(int a=0; a<args.Length; a++)
			{
				string curr = args[a];
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
			return true;
		}

		public static bool Recurse = false;
		public static bool Verbose = false;
		public static bool ShowAll = false;
		public static string Target = null;
	}
}