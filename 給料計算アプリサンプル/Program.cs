using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 給料計算アプリサンプル
{
	internal static class Program
	{
		/// <summary>
		/// アプリケーションのメイン エントリ ポイントです。
		/// </summary>
			[STAThread]
		static void Main()
		{
			CultureInfo culture = new CultureInfo("id-ID");
			Thread.CurrentThread.CurrentCulture = culture;
			Thread.CurrentThread.CurrentUICulture = culture;
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new Form1());
		}
	}
}
