using Npgsql;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections;
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
	public partial class Item_08 : Userbase
	{
		public Item_08()
		{
			InitializeComponent();
		}

		List<Work_time> wtList = new List<Work_time>();
		private void Item_08_Load(object sender, EventArgs e)
		{
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			con.Open();
			try
			{
				string sql = "SELECT wt_code,wt_shift  FROM work_time;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					wtList.Add(new Work_time(
						int.Parse(reader["wt_code"].ToString()),
						reader["wt_shift"].ToString()
					));
				}
				comboBox1.DataSource = wtList;
				comboBox1.ValueMember = "wt_code";
				comboBox1.DisplayMember = "shift";
			}
			finally
			{
				con.Close();
			}
			checkBox1.Checked = true;
			renewal_sh();
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
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan." + ex.Message);
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
			dt1.Columns.Add("Type");
			dt1.Columns.Add("Start Date");
			dt1.Columns.Add("Entry Date");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);

			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT shift.S_code,shift.e_code, personal_info.p_name, shift.S_type, shift.S_startdate, shift.S_entrydate " +
					"FROM shift left join personal_info on shift.e_code=personal_info.p_code";
				if(checkBox1.Checked)
				{
					sql = sql + " WHERE S_startdate <= CURRENT_DATE ORDER BY S_startdate DESC LIMIT 1;";
				}

				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["S_code"],
						reader["e_code"],
						reader["p_name"],
						reader["S_type"],
						reader["S_startdate"],
						reader["S_entrydate"]
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
		string excelFile;

		private void button2_Click(object sender, EventArgs e)
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
						string sql = "INSERT INTO shift(e_code, s_type, s_startdate, s_entrydatet)" +
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

		private void button1_Click(object sender, EventArgs e)
		{
			//新規単独登録

			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//Codeが選択されていません。
				MessageBox.Show("Code Karyawan tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox1.Text))
			{
				//残業時間が入力されていません。
				MessageBox.Show("JTipe tidak dipilih..");
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
						string query = "insert into shift(e_code,s_type,s_startdate, s_entrydate)" +
							" values(@code, @type, @start, @entry)";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox2.Text));
							cmd.Parameters.AddWithValue("@type", int.Parse(comboBox1.SelectedValue.ToString()));
							cmd.Parameters.AddWithValue("@start", dateTimePicker1.Value);
							cmd.Parameters.AddWithValue("@entry", DateTime.Now);

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
			renewal_sh();
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
					kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "?", "*perhatian*", MessageBoxButtons.YesNo);
					if (kekka == DialogResult.Yes)
					{
						con.Open();
						try
						{
							string query = "update shift set e_code=@code, s_type=@type, s_startdate=@start, " +
								"s_entrydate=@entry where s_code=@s_code";

							using (var cmd = new NpgsqlCommand(query, con))
							{
								// パラメータを設定
								cmd.Parameters.AddWithValue("@s_code", int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()));
								cmd.Parameters.AddWithValue("@code", int.Parse(dataGridView1.CurrentRow.Cells[1].Value.ToString()));
								cmd.Parameters.AddWithValue("@type", int.Parse(dataGridView1.CurrentRow.Cells[3].Value.ToString()));
								cmd.Parameters.AddWithValue("@start", DateTime.Parse(dataGridView1.CurrentRow.Cells[4].Value.ToString()));
								cmd.Parameters.AddWithValue("@entry", DateTime.Now);

								// クエリを実行
								cmd.ExecuteNonQuery();
							}

							renewal_sh();
							//登録に成功しました。
							MessageBox.Show("Pendaftaran berhasil.");
						}
						catch
						{ }
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
		List<P_Info> piList = new List<P_Info>();
		private void textBox2_Leave(object sender, EventArgs e)
		{
			try
			{
				int inputCode;
				if (int.TryParse(textBox2.Text, out inputCode))
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

		
	}
}
