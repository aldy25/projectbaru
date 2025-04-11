using Npgsql;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace 給料計算アプリサンプル
{
	
	public partial class Item_05 : Userbase
	{
		public Item_05()
		{
			InitializeComponent();
		}

		private void Item_05_Load(object sender, EventArgs e)
		{
			currentYear = DateTime.Now.Year;
			currentMonth = DateTime.Now.Month;
			makeCalender();
			renewal_sh();
		}
		int currentYear;
		int currentMonth;
		System.Windows.Forms.Button[] dayButtons = new System.Windows.Forms.Button[42];
		private void makeCalender()
		{
			int ButtonSize = 50;
			int StartX = 26, StartY = 122;
			DateTime firstDay = new DateTime(currentYear, currentMonth, 1);
			int startDayIndex = ((int)firstDay.DayOfWeek + 6) % 7; // 月曜日を0に調整
			int daysInMonth = DateTime.DaysInMonth(currentYear, currentMonth);
			label2.Text = firstDay.ToString("MMMM yyyy");
			// **1. すでにあるボタンを削除**
			try
			{
				foreach (System.Windows.Forms.Button btn in dayButtons)
				{
					if (btn != null)
					{
						this.Controls.Remove(btn);
						btn.Dispose(); // メモリ解放
					}
				}
			}
			catch { }
			for (int i = 0; i < dayButtons.Length; i++)
			{
				System.Windows.Forms.Button btn = new System.Windows.Forms.Button
				{
					Size = new Size(ButtonSize, ButtonSize),
					Location = new Point(StartX + ((ButtonSize + 6) * (i % 7)), StartY + ((ButtonSize + 6) * (i / 7))),
					BackColor = Color.LightGray,
					Tag = i
				};
				btn.Click += DayButton_Click;
				dayButtons[i] = btn;
				this.Controls.Add(btn);
			}
			

			// 日付をボタンに設定
			for (int i = 0; i < daysInMonth; i++)
			{
				int index = startDayIndex + i;
				dayButtons[index].Text = (i + 1).ToString();
				dayButtons[index].Enabled = true;
				// 土曜日 (index % 7 == 6)、日曜日 (index % 7 == 0) の色を変更
				if (index % 7 == 5) // 土曜日
				{
					dayButtons[index].BackColor = Color.LightCoral;
				}
				else if (index % 7 == 6) // 日曜日
				{
					dayButtons[index].BackColor = Color.LightCoral;
				}
			}

		}
		private void DayButton_Click(object sender, EventArgs e)
		{
			System.Windows.Forms.Button clickedButton = sender as System.Windows.Forms.Button;
			if (clickedButton != null)
			{
				clickedButton.BackColor = clickedButton.BackColor == Color.LightGray ? Color.LightCoral : Color.LightGray;
			}
		}

		private void button13_Click(object sender, EventArgs e)
		{

			foreach (System.Windows.Forms.Button btn in dayButtons)
			{
				if (btn.Enabled && btn.BackColor == Color.LightCoral) // 選択されたボタン
				{
					// Text に数字が含まれているか確認（正規表現を使用）
					if (Regex.IsMatch(btn.Text, @"\d"))
					{
						// dayString を int に変換
						int day = int.Parse(btn.Text); // または Convert.ToInt32(dayString)
														// DateTime インスタンスを作成
						DateTime date = new DateTime(currentYear, currentMonth, day);
						try
						{
							using (var con = new NpgsqlConnection(connectionString))
							{
								string query = "insert into Working_days(WD_date, WD_tipe) values(@date,@tipe) ON CONFLICT (WD_date) DO UPDATE SET WD_tipe = EXCLUDED.WD_tipe;";
								con.Open(); // ★ 接続を開く
								using (var cmd = new NpgsqlCommand(query, con))
								{
									// パラメータを設定
									cmd.Parameters.AddWithValue("@date", date);
									cmd.Parameters.AddWithValue("@tipe", 0);

									// クエリを実行
									cmd.ExecuteNonQuery();
								}
							}
						}
						catch { }
					}
				}
				else
				{
					// Text に数字が含まれているか確認（正規表現を使用）
					if (Regex.IsMatch(btn.Text, @"\d"))
					{
						// dayString を int に変換
						int day = int.Parse(btn.Text); // または Convert.ToInt32(dayString)
														// DateTime インスタンスを作成
						DateTime date = new DateTime(currentYear, currentMonth, day);
						try
						{
							string query = "insert into Working_days(WD_date, WD_tipe) values(@date,@tipe) ON CONFLICT (WD_date) DO UPDATE SET WD_tipe = EXCLUDED.WD_tipe;;";
							using (var con = new NpgsqlConnection(connectionString))
							{
								con.Open(); // ★ 接続を開く
								using (var cmd = new NpgsqlCommand(query, con))
								{
									// パラメータを設定
									cmd.Parameters.AddWithValue("@date", date);
									cmd.Parameters.AddWithValue("@tipe", 1);

									// クエリを実行
									cmd.ExecuteNonQuery();
								}
							}
						}
						catch { }
					}
				}
			}
			renewal_sh();
		}

		private void label1_Click(object sender, EventArgs e)
		{
			//前の月
			currentMonth--;
			if (currentMonth == 0)
			{
				currentYear--;
				currentMonth = 12;
				makeCalender();
				renewal_sh();
			}
			else
			{
				makeCalender();
				renewal_sh();
			}
		}

		private void label3_Click(object sender, EventArgs e)
		{
			//翌月
			currentMonth++;
			if (currentMonth == 12)
			{
				currentYear++;
				currentMonth = 1;
				makeCalender();
				renewal_sh();
			}
			else
			{
				makeCalender();
				renewal_sh();
			}
		}
		DataTable dt1;

		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Date");
			dt1.Columns.Add("Tipe");

			DateTime startDate = new DateTime(currentYear, currentMonth, 1);
			int lastDay = DateTime.DaysInMonth(currentYear, currentMonth);
			DateTime finalDate = new DateTime(currentYear, currentMonth, lastDay);

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);

			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM Working_days WHERE WD_date BETWEEN @startDate AND @endDate;";
				using (NpgsqlCommand com = new NpgsqlCommand(sql, con))
				{
					// パラメータの設定
					com.Parameters.AddWithValue("@startDate", startDate);
					com.Parameters.AddWithValue("@endDate", finalDate);

					using (NpgsqlDataReader reader = com.ExecuteReader())
					{
						while (reader.Read())
						{
							dt1.Rows.Add(
								reader["WD_date"],
								reader["WD_tipe"]
							);

							// WD_tipe のチェック（例: "Holiday" の場合）
							if (reader["WD_tipe"].ToString() == "0")
							{
								foreach (System.Windows.Forms.Button btn in dayButtons)
								{
									int wdDate = ((DateTime)reader["WD_date"]).Day;
									if (btn.Text==wdDate.ToString())
									{
										btn.BackColor = Color.LightCoral;
									}
								}
							}
							else if (reader["WD_tipe"].ToString() == "1")
							{
								foreach (System.Windows.Forms.Button btn in dayButtons)
								{
									int wdDate = ((DateTime)reader["WD_date"]).Day;
									if (btn.Text == wdDate.ToString())
									{
										btn.BackColor = SystemColors.Control;
									}
								}
							}
						}
					}
				}

				// DataGridViewにデータをバインド
				dataGridView1.DataSource = dt1;

				// DataGridViewのスタイル設定
				this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 14);
				dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
				dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

			}
			catch (Exception ex)
			{
				//問題が発生しリスト表示されませんでした。
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan.\n" + ex.Message);
			}
			finally
			{
				con.Close();
			}
		}
	}
}
