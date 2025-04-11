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
	public partial class Att_03 : Userbase
	{
		public Att_03()
		{
			InitializeComponent();
		}

		string excelFile;
		private void button4_Click(object sender, EventArgs e)
		{
			DialogResult kekka = openFileDialog1.ShowDialog();
			if (kekka == DialogResult.OK)
			{
				excelFile = openFileDialog1.FileName;
			}

			using (var con = new NpgsqlConnection(connectionString))
			{
				con.Open();
				// Excelデータをbasic_salaryに登録
				NpgsqlTransaction transaction = null;
				try
				{
					transaction = con.BeginTransaction();

					IWorkbook book = WorkbookFactory.Create(excelFile);
					ISheet sheet = book.GetSheet("Sheet1");

					for (int i = 0; i <= sheet.LastRowNum; i++)
					{
						IRow row = sheet.GetRow(i);
						string sql = "INSERT INTO overtime_schedule(e_code, os_date, os_hour, os_entrydatet)" +
							" VALUES";
						sql += "(";

						for (int j = 0; j < 4; j++)
						{
							ICell cell = row.GetCell(j);
							string str = "";

							if (cell != null)
							{
								switch (cell.CellType)
								{
									case CellType.Numeric:
										str = cell.NumericCellValue.ToString();
										break;
									case CellType.String:
										str = cell.StringCellValue.ToString();
										break;
								}
							}

							if (j != 0) sql += ",";
							sql += $"'{str.Replace("'", "''")}'"; // SQLインジェクション対策
						}

						sql += $");";
						using (var cmd = new NpgsqlCommand(sql, con, transaction))
						{
							cmd.ExecuteNonQuery();
						}
					}

					transaction.Commit();
					MessageBox.Show("データの読み込みに成功しました。");
				}
				catch
				{
					transaction?.Rollback();
					MessageBox.Show("データの読み込みに失敗しました。");
				}
			}
			renewal_sh();
		}

		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("code");
			dt1.Columns.Add("code karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Dept");
			dt1.Columns.Add("Date");
			dt1.Columns.Add("Hour");
			dt1.Columns.Add("entrydate");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT overtime_schedule.os_code, overtime_schedule.e_code, personal_info.p_name, personal_info.p_dept, " +
					"overtime_schedule.os_date, overtime_schedule.os_hour, overtime_schedule.os_entrydate " +
					"FROM basic_salary left join personal_info on overtime_schedule.e_code =personal_info.p_code";
				
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["os_code"],
						reader["e_code"],
						reader["p_name"],
						reader["p_dept"],
						reader["os_date"],
						reader["os_hour"],
						reader["os_entrydate"]
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

		private void button2_Click(object sender, EventArgs e)
		{
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//Codeが選択されていません。
				MessageBox.Show("Code Karyawan tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox1.Text))
			{
				//残業時間が入力されていません。
				MessageBox.Show("Jam lembur belum dimasukkan.");
				return;
			}

			using (var con = new NpgsqlConnection(connectionString))
			{
				
				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + textBox3.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "insert into overtime_schedule(e_code,os_date,os_hour, os_entrydate)" +
							" values(@code, @date, @hour, @entrydate)";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox2.Text));
							cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value);
							cmd.Parameters.AddWithValue("@hour", int.Parse(textBox1.Text));
							cmd.Parameters.AddWithValue("@entrydate", DateTime.Now);

							// クエリを実行
							cmd.ExecuteNonQuery();
						}
					}
					finally
					{
						con.Close();
					}
					renewal_sh();
					//登録に成功しました。
					MessageBox.Show("Pendaftaran berhasil.");
					
				}
			}
		}

		private void button3_Click(object sender, EventArgs e)
		{
			//データ変更
			try
			{
				using (var con = new NpgsqlConnection(connectionString))
				{
					string today = DateTime.Today.ToString();

					DialogResult kekka;
					//textBox2の内容を登録しますか？
					kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "?", "*perhatian*", MessageBoxButtons.YesNo);
					if (kekka == DialogResult.Yes)
					{
						con.Open();
						try
						{
							string query = "update overtime_schedule set e_code=@code, os_date=@date, os_hour=@hour" +
								"os_entrydate=@entry where os_code=@os";

							using (var cmd = new NpgsqlCommand(query, con))
							{
								// パラメータを設定
								cmd.Parameters.AddWithValue("@os", int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()));
								cmd.Parameters.AddWithValue("@code", int.Parse(dataGridView1.CurrentRow.Cells[1].Value.ToString()));
								cmd.Parameters.AddWithValue("@date", DateTime.Parse(dataGridView1.CurrentRow.Cells[4].Value.ToString()));
								cmd.Parameters.AddWithValue("@hour", int.Parse(dataGridView1.CurrentRow.Cells[5].Value.ToString()));
								cmd.Parameters.AddWithValue("@entry", DateTime.Now);

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
			catch
			{
				return;
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//データ削除
			try
			{
				using (var con = new NpgsqlConnection(connectionString))
				{
					string today = DateTime.Today.ToString();

					DialogResult kekka;
					//選択されている内容を削除しますか？
					kekka = MessageBox.Show(this, "Apakah Anda ingin menghapus konten yang " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "?", "*perhatian*", MessageBoxButtons.YesNo);
					if (kekka == DialogResult.Yes)
					{
						con.Open();
						try
						{
							string query = "delete from overtime_schedule where os_code=@os";

							using (var cmd = new NpgsqlCommand(query, con))
							{
								// パラメータを設定
								cmd.Parameters.AddWithValue("@os", int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()));

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
			catch
			{
				return;
			}
		}

		private void Att_03_Load(object sender, EventArgs e)
		{

		}
	}
}
