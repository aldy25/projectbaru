using Npgsql;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 給料計算アプリサンプル
{
	internal class ReBase
	{
		public string connectionString = ConfigurationManager.AppSettings["connectionString"];
		//従業員給料持ち越し用
		public class Payroll
		{
			public int P_code { get; set; }
			public string Name { get; set; }
			public string Dept { get; set; }
			public string Title { get; set; }
			public string Bank { get; set; }
			public long Basic { get; set; }
			public long Allowance { get; set; }
			public long Transport { get; set; }
			public long Absence { get; set; }
			public long Other { get; set; }
		    public long BPJSkesh { get; set; }
			public long BPJSNakar { get; set; }
			public long BPJSpen { get; set; }
			public long SPcost { get; set; }
			public double OTA { get; set; }
			public double OTB { get; set; }
			public double OTC { get; set; }
			public double OTD { get; set; }


			public Payroll(int p_code, string name, string dept, string title, string bank, long basic, long allowance, long transport, long absence, long other, long bpjskesh, long bpjsnakar, long bpjspen,
				long spcost, double ota, double otb, double otc, double otd)
			{
				P_code = p_code;
				Name = name;
				Dept = dept;
				Title = title;
				Bank = bank;
				Basic = basic;
				Allowance = allowance;
				Transport = transport;
				Absence = absence;
				Other = other;
				BPJSkesh = bpjskesh;
				BPJSNakar = bpjsnakar;
				BPJSpen = bpjspen;
				SPcost = spcost;
				OTA = ota;
				OTB = otb;
				OTC = otc;
				OTD = otd;
			}
		}
		public void ExecuteNonQuery(string query)
		{
			try
			{
				// 接続を作成
				using (var conn = new NpgsqlConnection(connectionString))
				{
					conn.Open(); // 接続を開く

					// コマンドを作成して実行
					using (var command = new NpgsqlCommand(query, conn))
					{
						command.ExecuteNonQuery();
					}
				}
			}
			catch (Exception ex)
			{
				// 例外が発生した場合はメッセージボックスを表示
				MessageBox.Show(ex.Message);
			}
		}
		public string FoldeNameGet()
		{
			var dirpath = Path.GetDirectoryName(Application.ExecutablePath);
			return (dirpath);
		}
	}
}
