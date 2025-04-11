using Npgsql;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.Formula.Functions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using NPOI.Util;

namespace 給料計算アプリサンプル
{
	public partial class Att_01 : Userbase
	{
		public Att_01()
		{
			InitializeComponent();
		}
		DataTable dataTable;
		private void button1_Click(object sender, EventArgs e)
		{
			string filePath;
			OpenFileDialog openFileDialog = new OpenFileDialog
			{
				Filter = "CSVファイル (*.csv)|*.csv",
				Title = "CSVファイルを選択してください"
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				filePath=openFileDialog.FileName;
			}
			else
			{
				return;
			}
			try
			{
				dataTable = new DataTable();
				string[] lines = File.ReadAllLines(filePath);

				if (lines.Length > 0)
				{
					// ヘッダーを読み取る
					string[] headers = lines[0].Split(';'); // セミコロンで分割
					foreach (string header in headers)
					{
						dataTable.Columns.Add(header.Trim());
					}

					// データを読み取る
					for (int i = 1; i < lines.Length; i++)
					{
						string[] data = lines[i].Split(';'); // セミコロンで分割
						dataTable.Rows.Add(data);
					}
				}

				dataGridView1.DataSource = dataTable;
			}
			catch (Exception ex)
			{
				MessageBox.Show("エラーが発生しました: " + ex.Message);
			}
		}


		private void button2_Click(object sender, EventArgs e)
		{
			try
			{
				NpgsqlConnection con = new NpgsqlConnection(connectionString);
				for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
				{
					//日勤か夜勤か判断する必要がある。
					int Shift;
					TimeSpan finishtime=TimeSpan.Zero;
					// データ取得クエリ
					con.Open();
					string sql = "SELECT shift.s_type, work_time.wt_finishtime FROM shift left join work_time on shift.s_type=work_time.wt_code" +
						" where e_code="+dataGridView1.Rows[i].Cells[1].Value.ToString() +";";
					NpgsqlCommand com = new NpgsqlCommand(sql, con);
					NpgsqlDataReader reader = com.ExecuteReader();
					while (reader.Read())
					{
						Shift = int.Parse(reader["s_type"].ToString());
						finishtime = TimeSpan.Parse(reader["wt_finishtime"].ToString());
					}
					reader.Close();
					TimeSpan starttime = new TimeSpan(7, 30, 0);
					if (finishtime != null && finishtime < starttime)
					{
						// 夜勤の場合
					}
					else
					{
						//日勤の場合
						//最小時刻と最大時刻を取得
						string[] timeArray =
							{	dataGridView1.Rows[i].Cells[4].Value.ToString(),
								dataGridView1.Rows[i].Cells[5].Value.ToString(),
								dataGridView1.Rows[i].Cells[6].Value.ToString(),
								dataGridView1.Rows[i].Cells[7].Value.ToString(),
								dataGridView1.Rows[i].Cells[8].Value.ToString(),
								dataGridView1.Rows[i].Cells[9].Value.ToString(),
								dataGridView1.Rows[i].Cells[10].Value.ToString(),
								dataGridView1.Rows[i].Cells[11].Value.ToString()
					};
						// 0:00 を除外
						var filteredTimes = timeArray.Where(time => time != "00:00").ToList();
						if (filteredTimes.Count != 0) //すべてが0:00だった場合を回避
						{
							// 時刻を TimeSpan 型に変換
							var timeSpans = filteredTimes.Select(time => TimeSpan.Parse(time)).ToList();

							// 最小時刻と最大時刻を取得
							var minTime = timeSpans.Min();
							var maxTime = timeSpans.Max();

							if (dataGridView1.Rows[i].Cells[4].Value.ToString() != null && dataGridView1.Rows[i].Cells[4].Value.ToString() != "00:00")
							{
								//string dateOnly = dateTimePicker1.Value.ToString("dd/MM/yyyy");
								//if (dataGridView1.Rows[i].Cells[0].Value.ToString() == dateOnly)
								//{

									using (NpgsqlConnection conn = new NpgsqlConnection(connectionString))
									{
										conn.Open();
										try
										{
											string query = "insert into working_status(e_code, w_date, w_starttime, w_finishtime) " +
												"values(@code, @date, @starttime, @finishtime);";

											using (var cmd = new NpgsqlCommand(query, conn))
											{
												// パラメータを設定
												cmd.Parameters.AddWithValue("@code", int.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString()));
												cmd.Parameters.AddWithValue("@date", DateTime.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString()));
												cmd.Parameters.AddWithValue("@starttime", minTime);
												cmd.Parameters.AddWithValue("@finishtime", maxTime);

												// クエリを実行
												cmd.ExecuteNonQuery();
											}


										}
										catch { MessageBox.Show("失敗"); }
										finally
										{
											conn.Close();
										}
									}
								//}
							}
						}
						con.Close();
					}
				}
			}
			catch
			{ }
			finally
			{
				//登録に成功しました。
				MessageBox.Show("Pendaftaran berhasil.");
				dataTable = new DataTable();
				dataGridView1.DataSource = dataTable;
			}
		}

		private void button3_Click(object sender, EventArgs e)
		{
			dataTable = new DataTable();
			dataGridView1.DataSource = dataTable;
		}

		private void Att_01_Load(object sender, EventArgs e)
		{

		}
	}
}
