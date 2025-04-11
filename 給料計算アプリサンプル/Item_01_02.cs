using Npgsql;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static 給料計算アプリサンプル.Userbase;

namespace 給料計算アプリサンプル
{
	public partial class Item_01_02 : Form
	{
		public Item_01_02()
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
		private void button4_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		public string connectionString = ConfigurationManager.AppSettings["connectionString"];
		List<Dept> deptList = new List<Dept>();  //combobox1
		List<Title> titleList = new List<Title>();

		private void Item_01_02_Load(object sender, EventArgs e)
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

				comboBox3.Items.Add("Bekerja");
				comboBox3.Items.Add("Pensiun");
				comboBox3.SelectedIndex = 0;
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
		private void button3_Click(object sender, EventArgs e)
		{
			DialogResult kekka = openFileDialog1.ShowDialog();
			if (kekka == DialogResult.OK)
			{
				excelFile = openFileDialog1.FileName;
			}
			string dbpass = ConfigurationManager.AppSettings["connectionString"];
			using (var con = new NpgsqlConnection(dbpass))
			{
				con.Open();
				// Excelデータをcatalogに登録
				NpgsqlTransaction transaction = null;
				try
				{
					transaction = con.BeginTransaction();

					IWorkbook book = WorkbookFactory.Create(excelFile);
					ISheet sheet = book.GetSheet("Sheet1");

					for (int i = 0; i <= sheet.LastRowNum; i++)
					{
						IRow row = sheet.GetRow(i);
						string sql = "INSERT INTO public_holiday VALUES";
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
		DataTable dt1;
		private void button1_Click(object sender, EventArgs e)
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code Karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Dept");
			dt1.Columns.Add("Title");
			dt1.Columns.Add("Mail");
			dt1.Columns.Add("Status");
			dt1.Columns.Add("SP");
			dt1.Columns.Add("Update");
			dt1.Columns.Add("StartDate");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			int i = 0;//,が必要かどうか判定
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM personal_info";
				if (!string.IsNullOrWhiteSpace(textBox1.Text))
				{
					i++;
					sql = sql + " where p_code=like %" + textBox1.Text + "%";
				}
				if (!string.IsNullOrWhiteSpace(textBox2.Text))
				{

					if (i == 0)
					{
						sql = sql + " where p_name=like %" + textBox2.Text + "% ";
					}
					else
					{
						sql = sql + " and p_name=like %" + textBox2.Text + "% ";
					}
					i++;
				}
				if (comboBox1.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox1.SelectedValue.ToString()))
				{

					if (i == 0)
					{
						sql = sql + " where p_dept=" + comboBox1.SelectedValue.ToString();
					}
					else
					{
						sql = sql + " and p_dept=" + comboBox1.SelectedValue.ToString();
					}
					i++;
				}
				if (comboBox2.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox2.SelectedValue.ToString()))
				{

					if (i == 0)
					{
						sql = sql + " where p_title=" + comboBox2.SelectedValue.ToString();
					}
					else
					{
						sql = sql + " and p_dept=" + comboBox2.SelectedValue.ToString();
					}
					i++;
				}
				if (comboBox3.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox3.SelectedValue.ToString()))
				{

					if (i == 0)
					{
						sql = sql + " where p_status=" + comboBox3.SelectedValue.ToString();
					}
					else
					{
						sql = sql + " and p_status=" + comboBox3.SelectedValue.ToString();
					}
					i++;
				}
				
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["p_code"],
						reader["p_name"],
						reader["p_dept"],
						reader["p_title"],
						reader["p_mail"],
						reader["p_status"],
						reader["p_sp"],
						reader["p_update"],
						reader["p_startdate"]
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
			using (var con = new NpgsqlConnection(connectionString))
			{
				string today = DateTime.Today.ToString();

				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan apa yang Anda lihat di daftar?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						for (int i = 0; i < dataGridView1.Rows.Count; i++)
						{
							string query = "insert into personal_info(p_code, p_name, p_dept, p_title, p_mail, p_status, p_sp, p_update, p_startdate) " +
								"values(@code, @name, @dept, @title, @mail, @status, @sp, @update, @start);";

							using (var cmd = new NpgsqlCommand(query, con))
							{
								// パラメータを設定
								cmd.Parameters.AddWithValue("@code", dataGridView1.Rows[i].Cells[0].Value.ToString());
								cmd.Parameters.AddWithValue("@name", dataGridView1.Rows[i].Cells[1].Value.ToString());
								cmd.Parameters.AddWithValue("@dept", dataGridView1.Rows[i].Cells[2].Value.ToString());
								cmd.Parameters.AddWithValue("@title", dataGridView1.Rows[i].Cells[3].Value.ToString());
								cmd.Parameters.AddWithValue("@mail", dataGridView1.Rows[i].Cells[4].Value.ToString());
								cmd.Parameters.AddWithValue("@status", dataGridView1.Rows[i].Cells[5].Value.ToString());
								cmd.Parameters.AddWithValue("@sp", dataGridView1.Rows[i].Cells[6].Value.ToString());
								cmd.Parameters.AddWithValue("@update", dataGridView1.Rows[i].Cells[7].Value.ToString());
								cmd.Parameters.AddWithValue("@start", dataGridView1.Rows[i].Cells[8].Value.ToString());
								// クエリを実行
								cmd.ExecuteNonQuery();
							}
						}
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
	}
}
