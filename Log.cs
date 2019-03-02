using System;

namespace LnkUtils
{
	public static class Log
	{
		public static void Message(string m)
		{
			Console.WriteLine(m);
		}
		public static void Error(string m)
		{
			Console.Error.WriteLine("E: "+m);
		}
		public static void Warning(string m)
		{
			Console.WriteLine("W: "+m);
		}
		public static void Debug(string m)
		{
			#if DEBUG
			Console.Error.WriteLine("D: "+m);
			#endif
		}
	}
}