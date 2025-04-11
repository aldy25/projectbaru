using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace 給料計算アプリサンプル
{
	public partial class Userbase : UserControl
	{
		public string connectionString = ConfigurationManager.AppSettings["connectionString"];
		public Userbase()
		{
			InitializeComponent();
		}
		public class Dept
		{
			public int DP_code { get; set; }
			public string DP_name { get; set; }
			public Dept(int dp_code, string dp_name)
			{
				DP_code = dp_code;
				DP_name = dp_name;
			}
		}
		public class Title
		{
			public int T_code { get; set; }
			public string T_name { get; set; }
			public long T_allowance { get; set; }
			public Title(int t_code, string t_name, long t_allowance)
			{
				T_code = t_code;
				T_name = t_name;
				T_allowance = t_allowance;
			}
		}
		public class BS
		{
			public int B_code { get; set; }
			public string E_code { get; set; }
			public long B_salary { get; set; }
			public long B_allowance { get; set; }
			public long B_transport { get; set; }
			public long B_startdate { get; set; }
			public string B_update { get; set; }
			public BS(int b_code, string e_code, long b_salary, long b_allowance,long b_transport, long b_startdate, string b_update)
			{
				B_code = b_code;
				E_code = e_code;
				B_salary = b_salary;
				B_allowance = b_allowance;
				B_transport = b_transport;
				B_startdate = b_startdate;
				B_update = b_update;
			}
		}
		public class P_Info
		{
			public int PI_code { get; set; }
			public string PI_name { get; set; }
			public string PI_Dept { get; set; }
			public P_Info(int pi_code, string pi_name, string pi_dept)
			{
				PI_code = pi_code;
				PI_name = pi_name;
				PI_Dept = pi_dept;
			}
		}
		public class Work_time
		{
			public int Wt_code { get; set; }
			public string Shift { get; set; }
			public Work_time(int wt_code, string shift)
			{
				Wt_code = wt_code;
				Shift = shift;
			}
		}

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
			public long Rapelan { get; set; }
			public long Adj { get; set; }
			public long Other { get; set; }
			public long Absence { get; set; }
			public long BPJSkesh { get; set; }
			public long BPJSNakar { get; set; }
			public long BPJSpen　{ get; set; }
			public long SPcost　{ get; set; }
			public double OTA 　{ get; set; }
			public double OTB 　{ get; set; }
			public double OTC 　{ get; set; }
			public double OTD 　{ get; set; }
			
			public Payroll(int p_code, string name, string dept, string title, string bank, long basic, long allowance, long transport, long rapelan, long adj,long absence, long other,long bpjskesh,long bpjsnakar, long bpjspen, 
				long spcost, double ota, double otb, double otc, double otd)
			{
				P_code = p_code;
				Name = name;
				Dept = dept;
				Title = title;
				Bank = bank;
				Basic = basic;
				Allowance=allowance;
				Transport = transport;
				Rapelan = rapelan;
				Adj = adj;
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

		private void Userbase_Load(object sender, EventArgs e)
		{

		}
	}
}
