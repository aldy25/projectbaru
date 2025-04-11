using Npgsql;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace 給料計算アプリサンプル
{
	public partial class Att_04 : Userbase
	{
		public Att_04()
		{
			InitializeComponent();
		}

		private void Att_04_Load(object sender, EventArgs e)
		{
			renewal_sh();
		}

		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code Karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Dept");
			dt1.Columns.Add("Shift");
			dt1.Columns.Add("Status");//出退勤ステータス、出勤しているかどうか？
			dt1.Columns.Add("Waktu masuk kerja");
			dt1.Columns.Add("Cuti");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			int status=2;//出社0, 欠席1, 無断欠席2
			int MK = 0;//出社人数
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT pi.p_code,pi.p_name,pi.p_dept,s.s_type,ws.w_starttime,d.d_tipe"+
					" FROM personal_info AS pi JOIN shift AS s ON pi.p_code = s.e_code"+
					" AND s.s_startdate<CURRENT_DATE"+
					" JOIN working_status AS ws ON pi.p_code = ws.e_code"+
					" JOIN day_off AS d ON pi.p_code = d.e_code"+
					" AND ws.w_date = d.d_date WHERE d.d_date = '"+dateTimePicker1.Value.ToString("yyyy-MM-dd") +"'";
				
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					//statusのデータを作成する
					if(reader["s_type"].ToString()=="0")
					{
						if (DateTime.Parse(reader["w_starttime"].ToString()) > DateTime.Parse("5:00"))
						{
							status = 0;
							MK++;
						}
						else
						{
							if (reader["d_tipe"].ToString() != null && !string.IsNullOrWhiteSpace(reader["d_tipe"].ToString()))
							{
								status = 1;
							}
						}
					}
					dt1.Rows.Add(
						reader["p_code"],
						reader["p_name"],
						reader["p_dept"],
						reader["s_type"],
						status,
						reader["w_starttime"],
						reader["d_tipe"]
					);
				}

				// DataGridViewにデータをバインド
				dataGridView1.DataSource = dt1;

				// DataGridViewのスタイル設定
				this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 14);
				dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
				dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

				reader.Close();

				textBox1.Text=MK.ToString();
			}
			catch (Exception ex)
			{
				//問題が発生しリスト表示されませんでした。
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan." + ex.Message);
			}
			finally
			{
				con.Close();
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			//whatsapp自動メール送信　出席を確定する
			MessageBox.Show("belum selasai");
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//whatsapp自動メール送信無断欠席について担当者に連絡する
			MessageBox.Show("belum selasai");
		}
	}
}
