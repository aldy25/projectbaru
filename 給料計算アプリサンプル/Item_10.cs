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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace 給料計算アプリサンプル
{
	public partial class Item_10 : Userbase
	{
		public Item_10()
		{
			InitializeComponent();
		}

		private void Item_10_Load(object sender, EventArgs e)
		{
			renewal_sh();

			//dateTimepickerすべてに0:00を入力
			for (int i = 1; i <= 15; i++)
			{
				// コントロール名を動的に作成
				string controlName = $"datetimepicker{i}";

				// コントロールを取得
				var dateTimePicker = this.Controls.Find(controlName, true).FirstOrDefault() as DateTimePicker;

				if (dateTimePicker != null)
				{
					// 時刻を 0:00 に設定
					dateTimePicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
				}
			}

		}

		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code");
			dt1.Columns.Add("Name Shift");
			dt1.Columns.Add("Mulai Bekerja");
			dt1.Columns.Add("Akhir jam kerja");
			dt1.Columns.Add("Akhir jam kerja OT");
			dt1.Columns.Add("Istirahat1 mulai");
			dt1.Columns.Add("Istirahat1 selasai");
			dt1.Columns.Add("Istirahat2 mulai");
			dt1.Columns.Add("Istirahat2 selasai");
			dt1.Columns.Add("Istirahat3 mulai");
			dt1.Columns.Add("Istirahat3 selasai");
			dt1.Columns.Add("Istirahat4 mulai");
			dt1.Columns.Add("Istirahat4 selasai");
			dt1.Columns.Add("Istirahat5 mulai");
			dt1.Columns.Add("Istirahat5 selasai");
			dt1.Columns.Add("Istirahat6 mulai");
			dt1.Columns.Add("Istirahat6 selasai");
			dt1.Columns.Add("tanggal masukan");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM work_time";

				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["wt_code"],
						reader["wt_shift"],
						reader["wt_starttime"],
						reader["wt_finishtime"],
						reader["wt_otftime"],
						reader["wt_break1"],
						reader["wt_break2"],
						reader["wt_break3"],
						reader["wt_break4"],
						reader["wt_break5"],
						reader["wt_break6"],
						reader["wt_break7"],
						reader["wt_break8"],
						reader["wt_break9"],
						reader["wt_break10"],
						reader["wt_break11"],
						reader["wt_break12"],
						reader["wt_entrydate"]
					);
				}

				// DataGridViewにデータをバインド
				dataGridView1.DataSource = dt1;

				// DataGridViewのスタイル設定
				this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 14);
				dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
				dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

				reader.Close();
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


		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
			textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
			//dateTimepickerに入力
			for (int i = 1; i <= 15; i++)
			{
				// コントロール名を動的に作成
				string controlName = $"datetimepicker{i}";

				// コントロールを取得
				var dateTimePicker = this.Controls.Find(controlName, true).FirstOrDefault() as DateTimePicker;
				try
				{
					if (dateTimePicker != null)
					{
						// 時刻を 0:00 に設定
						dateTimePicker.Value = DateTime.Parse(dataGridView1.CurrentRow.Cells[i + 1].Value.ToString());
					}
				}
				catch 
				{
					dateTimePicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
				}
			}
			
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox1.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Kode tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//シフトの名前が入力されていません。
				MessageBox.Show("Nama shift belum dimasukkan.");
				return;
			}
			
			using (var con = new NpgsqlConnection(connectionString))
			{
				string today = DateTime.Today.ToString();

				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + textBox2.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "update work_time set wt_shifte=@name, wt_starttime=@starttime, wt_finishtime, wt_otftime=@otftime" +
							"wt_break1=@break1,wt_break2=@break2, wt_break3=@break3,wt_break4=@break4, wt_break5=@break5, wt_break6=@break6" +
							"wt_break7=@break7, wt_break8=@break8, wt_break9=@break9, wt_break10=@break10, wt_break11=@break11, wt_break12=@break12" +
							", wt_entrydate=entrydate) " +
							"where code=@code;";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox1.Text));
							cmd.Parameters.AddWithValue("@name", textBox2.Text);
							cmd.Parameters.AddWithValue("@starttime", dateTimePicker1.Value);
							cmd.Parameters.AddWithValue("@finishtime", dateTimePicker2.Value);
							cmd.Parameters.AddWithValue("@otftime", dateTimePicker3.Value);
							cmd.Parameters.AddWithValue("@break1", dateTimePicker4.Value);
							cmd.Parameters.AddWithValue("@break2", dateTimePicker5.Value);
							cmd.Parameters.AddWithValue("@break3", dateTimePicker6.Value);
							cmd.Parameters.AddWithValue("@break4", dateTimePicker7.Value);
							cmd.Parameters.AddWithValue("@break5", dateTimePicker8.Value);
							cmd.Parameters.AddWithValue("@break6", dateTimePicker9.Value);
							cmd.Parameters.AddWithValue("@break7", dateTimePicker10.Value);
							cmd.Parameters.AddWithValue("@break8", dateTimePicker11.Value);
							cmd.Parameters.AddWithValue("@break9", dateTimePicker12.Value);
							cmd.Parameters.AddWithValue("@break10", dateTimePicker13.Value);
							cmd.Parameters.AddWithValue("@break11", dateTimePicker14.Value);
							cmd.Parameters.AddWithValue("@break12", dateTimePicker15.Value);
							cmd.Parameters.AddWithValue("@entrydatel", today);

							// クエリを実行
							cmd.ExecuteNonQuery();
						}

						renewal_sh();
						//登録に成功しました。
						MessageBox.Show("Pendaftaran berhasil.");
					}
					finally
					{
						con.Close();
					}
				}
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//シフトの名前が入力されていません。
				MessageBox.Show("Nama shift belum dimasukkan.");
				return;
			}
			using (var con = new NpgsqlConnection(connectionString))
			{
				DateTime today = DateTime.Today;

				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + textBox2.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "insert into work_time(wt_shift, wt_starttime, wt_finishtime, wt_otftime, wt_break1, wt_break2, wt_break3, wt_break4" +
							", wt_break5, wt_break6, wt_break7, wt_break8, wt_break9, wt_break10, wt_break11, wt_break12, wt_entrydate) " +
							"values(@name, @starttime, @finishtime, @otftime, @break1, @break2, @break3, @break4, @break5, @break6, @break7," +
							"@break8, @break9, @break10, @break11, @break12, @entrydate);";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@name", textBox2.Text);
							cmd.Parameters.AddWithValue("@starttime", dateTimePicker1.Value);
							cmd.Parameters.AddWithValue("@finishtime", dateTimePicker2.Value);
							cmd.Parameters.AddWithValue("@otftime", dateTimePicker3.Value);
							cmd.Parameters.AddWithValue("@break1", dateTimePicker4.Value);
							cmd.Parameters.AddWithValue("@break2", dateTimePicker5.Value);
							cmd.Parameters.AddWithValue("@break3", dateTimePicker6.Value);
							cmd.Parameters.AddWithValue("@break4", dateTimePicker7.Value);
							cmd.Parameters.AddWithValue("@break5", dateTimePicker8.Value);
							cmd.Parameters.AddWithValue("@break6", dateTimePicker9.Value);
							cmd.Parameters.AddWithValue("@break7", dateTimePicker10.Value);
							cmd.Parameters.AddWithValue("@break8", dateTimePicker11.Value);
							cmd.Parameters.AddWithValue("@break9", dateTimePicker12.Value);
							cmd.Parameters.AddWithValue("@break10", dateTimePicker13.Value);
							cmd.Parameters.AddWithValue("@break11", dateTimePicker14.Value);
							cmd.Parameters.AddWithValue("@break12", dateTimePicker15.Value);
							cmd.Parameters.AddWithValue("@entrydate", today);

							// クエリを実行
							cmd.ExecuteNonQuery();
						}

						renewal_sh();
						//登録に成功しました。
						MessageBox.Show("Pendaftaran berhasil.");
					}
					finally
					{
						con.Close();
					}
					renewal_sh();
				}
			}
		}
	}
}
