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
using static NPOI.HSSF.Util.HSSFColor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace 給料計算アプリサンプル
{
	public partial class Att_02_01 : Userbase
	{
		public Att_02_01()
		{
			InitializeComponent();
		}
		List<Dept> deptList = new List<Dept>();
		List<Title> titleList = new List<Title>();
		List<P_Info> piList = new List<P_Info>();

		private void Att_02_01_Load(object sender, EventArgs e)
		{
			//部門Departmentのデータをダウンロード
			//部門のデータをlistに入れてcombobox1とcombobox4に入れる。
			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);

			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM department;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					deptList.Add(new Dept(
						int.Parse(reader["dp_code"].ToString()),
						reader["dp_name"].ToString()
					));

				}
				comboBox1.DataSource = deptList;
				comboBox1.ValueMember = "dp_code";
				comboBox1.DisplayMember = "dp_name";
				comboBox1.SelectedIndex = -1;
				reader.Close();

				string sql2 = "SELECT * FROM title;";
				NpgsqlCommand com2 = new NpgsqlCommand(sql2, con);
				NpgsqlDataReader reader2 = com2.ExecuteReader();
				while (reader2.Read())
				{
					titleList.Add(new Title(
						int.Parse(reader2["t_code"].ToString()),
						reader["t_name"].ToString(),
						int.Parse(reader2["t_allowance"].ToString())
					));
				}
				comboBox2.DataSource = titleList;
				comboBox2.ValueMember = "t_code";
				comboBox2.DisplayMember = "t_name";
				comboBox2.SelectedIndex = -1;

				reader2.Close();
			}
			catch (Exception ex)
			{
				//問題が発生しリスト表示されませんでした。
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan. " + ex.Message);
			}
			finally
			{
				con.Close();
			}
			renewal_sh();

			comboBox3.Items.Add("Cuti");
			comboBox3.Items.Add("Cuti Sakit");
			comboBox3.Items.Add("liburan yang tidak dibayar");
			comboBox3.Items.Add("lain");

			//Personal_Infoの情報をリストに入れておく
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM personal_info;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					piList.Add(new P_Info(
						int.Parse(reader["p_code"].ToString()),
						reader["p_name"].ToString(),
						reader["p_dept"].ToString()
					));

				}
				reader.Close();
			}
			catch (Exception ex)
			{
				//問題が発生しリスト表示されませんでした。
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan. " + ex.Message);
			}
			finally
			{
				con.Close();
			}
		}
		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code");
			dt1.Columns.Add("Code Karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Dept");
			dt1.Columns.Add("Title");
			dt1.Columns.Add("Date");
			dt1.Columns.Add("Entry Date");
			dt1.Columns.Add("Type");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			int i = 0;//,が必要かどうか判定
			try
			{
				con.Open();
				string sql = "SELECT day_off.d_code, day_off.e_code, personal_info.p_name, personal_info.p_dept, personal_info.p_title, " +
							"day_off.d_date, day_off.d_entrydate, day_off.d_tipe " +
							"FROM day_off LEFT JOIN personal_info ON day_off.e_code = personal_info.p_code WHERE ";

				List<string> conditions = new List<string>();

				if (!string.IsNullOrWhiteSpace(textBox1.Text))
				{
					conditions.Add($"day_off.e_code LIKE '%{textBox1.Text}%'");
				}
				if (!string.IsNullOrWhiteSpace(textBox2.Text))
				{
					conditions.Add($"personal_info.p_name LIKE '%{textBox2.Text}%'");
				}
				if (comboBox1.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox1.SelectedValue.ToString()))
				{
					conditions.Add($"personal_info.p_dept = '{comboBox1.SelectedValue.ToString()}'");
				}
				if (comboBox2.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox2.SelectedValue.ToString()))
				{
					conditions.Add($"personal_info.p_title = '{comboBox2.SelectedValue.ToString()}'");
				}

				if (conditions.Count > 0)
				{
					sql += string.Join(" AND ", conditions) + " AND ";
				}
				sql += "personal_info.p_startdate <= CURRENT_DATE ORDER BY day_off.d_code LIMIT 300";


				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["d_code"],
						reader["e_code"],
						reader["p_name"],
						reader["p_dept"],
						reader["p_title"],
						reader["d_date"],
						reader["d_entrydate"],
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

		private void button1_Click(object sender, EventArgs e)
		{
			renewal_sh();
		}

		private void button2_Click(object sender, EventArgs e)
		{
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox4.Text))
			{
				//Codeが選択されていません。
				MessageBox.Show("Code tidak dipilih.");
				return;
			}

			using (var con = new NpgsqlConnection(connectionString))
			{
				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + textBox5.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "update day_off set d_date = @date, d_entrydate = @entry, d_tipe = @tipe " +
							"where d_code=@code;";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox4.Text));
							cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value);
							cmd.Parameters.AddWithValue("@entry", DateTime.Now);
							cmd.Parameters.AddWithValue("@tipe", comboBox3.SelectedIndex);
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



		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			try
			{
				textBox4.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
				textBox5.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
				textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
				dateTimePicker1.Value = DateTime.Parse(dataGridView1.CurrentRow.Cells[5].Value.ToString());
				comboBox3.SelectedIndex = int.Parse(dataGridView1.CurrentRow.Cells[7].Value.ToString());
			}
			catch { }
		}

		private void textBox4_Leave(object sender, EventArgs e)
		{
			try
			{
				int inputCode;
				if (int.TryParse(textBox5.Text, out inputCode))
				{
					// piListから一致するPI_codeを検索
					var matchedItem = piList.FirstOrDefault(p => p.PI_code == inputCode);

					if (matchedItem != null)
					{
						// textBox3にPI_nameを表示
						textBox3.Text = matchedItem.PI_name;
					}
					else
					{
						// 一致するデータが見つからない場合
						MessageBox.Show("一致するデータが見つかりません。");
					}
				}
				else
				{
					// 数値が入力されていない場合
					MessageBox.Show("有効な数値を入力してください。");
				}
			}
			catch { }
		}

		string excelFile;
		private void button3_Click(object sender, EventArgs e)
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
						string sql = "INSERT INTO day_off(e_code, d_date, d_entrydate, d_tipe)" +
							" VALUES";
						sql += "(";

						for (int j = 0; j < 2; j++)
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
		}

		private void button4_Click(object sender, EventArgs e)
		{
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox5.Text))
			{
				//Codeが選択されていません。
				MessageBox.Show("Code Karyawan tidak dipilih.");
				return;
			}

			using (var con = new NpgsqlConnection(connectionString))
			{
				if (dateTimePicker1.Value.Date > dateTimePicker2.Value.Date)
				{
					MessageBox.Show("Rentang tanggal tidak valid. Tanggal mulai harus sama dengan atau lebih awal dari tanggal akhir.",
									"kesalahan",
									MessageBoxButtons.OK,
									MessageBoxIcon.Error);
					return;
				}
				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + textBox3.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{

						DateTime startDate = dateTimePicker1.Value.Date;
						DateTime endDate = dateTimePicker2.Value.Date;

						// 開始日から終了日までの日付をループ
						for (DateTime currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddDays(1))
						{
							string query = "insert into day_off(e_code, d_date, d_entrydate, d_tipe) values(@code, @date, @entry, @tipe);";

							using (var cmd = new NpgsqlCommand(query, con))
							{
								// パラメータを設定
								cmd.Parameters.AddWithValue("@code", int.Parse(textBox5.Text));
								cmd.Parameters.AddWithValue("@date", currentDate);
								cmd.Parameters.AddWithValue("@entry", DateTime.Now);
								cmd.Parameters.AddWithValue("@tipe", comboBox3.SelectedIndex);

								// クエリを実行
								cmd.ExecuteNonQuery();
							}
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

		private void textBox5_TextChanged(object sender, EventArgs e)
		{

		}
	}
}
