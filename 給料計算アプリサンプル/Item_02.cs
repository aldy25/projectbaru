using Npgsql;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static NPOI.HSSF.Util.HSSFColor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace 給料計算アプリサンプル
{
	public partial class Item_02 : Userbase
	{
		public Item_02()
		{
			InitializeComponent();
		}
		List<Dept> deptList=new List<Dept>();  //combobox1
		List<Dept> deptList2 = new List<Dept>(); 

		List<Title> titleList=new List<Title>();
		List<Title> titleList2 = new List<Title>();



		private void Item_02_Load(object sender, EventArgs e)
		{

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

					deptList2.Add(new Dept(
						int.Parse(reader["dp_code"].ToString()),
						reader["dp_name"].ToString()
					));
				}
				comboBox1.DataSource = deptList;
				comboBox1.ValueMember = "dp_code";
				comboBox1.DisplayMember = "dp_name";
				comboBox1.SelectedIndex = -1;

				comboBox4.DataSource = deptList2;
				comboBox4.ValueMember = "dp_code";
				comboBox4.DisplayMember = "dp_name";
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
					titleList2.Add(new Title(
						int.Parse(reader2["t_code"].ToString()),
						reader["t_name"].ToString(),
						int.Parse(reader2["t_allowance"].ToString())
					));
				}
				comboBox2.DataSource = titleList;
				comboBox2.ValueMember = "t_code";
				comboBox2.DisplayMember = "t_name";
				comboBox2.SelectedIndex = -1;

				comboBox5.DataSource = titleList2;
				comboBox5.ValueMember = "t_code";
				comboBox5.DisplayMember = "t_name";
				reader2.Close();

				renewal_sh();
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
			
			dt1.Columns.Add("code karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Dept");
			dt1.Columns.Add("Title");
			dt1.Columns.Add("code");
			dt1.Columns.Add("salary");
			dt1.Columns.Add("allowance");
			dt1.Columns.Add("transport");
			dt1.Columns.Add("update");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			int i = 0;//,が必要かどうか判定
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = @"
			WITH ranked_personal_info AS (
				SELECT *, 
						ROW_NUMBER() OVER (PARTITION BY p_code ORDER BY p_startdate DESC) AS rn
				FROM personal_info
			)
            SELECT 
                pi.p_code, 
                pi.p_name, 
                pi.p_dept, 
                pi.p_title, 
				COALESCE(bs.b_code, 0) AS b_code,
                COALESCE(bs.b_salary, 0) AS b_salary, 
				COALESCE(bs.b_allowance, 0) AS b_allowance, 
				COALESCE(bs.b_transport, 0) AS b_transport, 
                bs.b_update
            FROM 
                ranked_personal_info pi
            left JOIN 
                basic_salary bs
            ON 
                pi.p_code = bs.e_code
            WHERE
				pi.rn = 1
                AND pi.p_startdate <= CURRENT_DATE 
                AND (bs.b_startdate <= CURRENT_DATE OR bs.b_startdate IS NULL)";
				if (!string.IsNullOrWhiteSpace(textBox1.Text))
				{
					sql += $" AND bs.e_code LIKE '%{textBox1.Text}%'";
					i++;
				}

				if (!string.IsNullOrWhiteSpace(textBox7.Text))
				{
					sql += $" AND pi.p_name LIKE '%{textBox7.Text}%'";
					i++;
				}

				if (comboBox1.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox1.SelectedValue.ToString()))
				{
					sql += $" AND pi.p_dept = {comboBox1.SelectedValue}";
					i++;
				}

				if (comboBox2.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox2.SelectedValue.ToString()))
				{
					sql += $" AND pi.p_title = {comboBox2.SelectedValue}";
					i++;
				}

				sql += " AND (bs.b_flag IS NULL OR bs.b_flag IS NULL) ORDER BY pi.p_update DESC NULLS LAST, bs.b_startdate DESC";
				//pi.p_code LIMIT 1;

				string doublePI_code="";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{

					if (doublePI_code != reader["p_code"].ToString())
					{
						doublePI_code = reader["p_code"].ToString();
						dt1.Rows.Add(
							reader["p_code"],
							reader["p_name"],
							reader["p_dept"],
							reader["p_title"],
							reader["b_code"],
							reader["b_salary"],
							reader["b_allowance"],
							reader["b_transport"],
							reader["b_update"]
						);
					}
				}

				// DataGridViewにデータをバインド
				dataGridView1.DataSource = dt1;

				// DataGridViewのスタイル設定
				//***従業員コード順に並べ替えする***未完
				this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 14);
				dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
				dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

				reader.Close();
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
		private void button1_Click(object sender, EventArgs e)
		{
			renewal_sh();
		}

		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			try
			{
				textBox2.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
				textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
				comboBox4.SelectedValue = int.Parse(dataGridView1.CurrentRow.Cells[2].Value.ToString());
				comboBox5.SelectedValue = int.Parse(dataGridView1.CurrentRow.Cells[3].Value.ToString());
				textBox4.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
				textBox5.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
				textBox6.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
				textBox8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
			}
			catch { }
		}

		private void button3_Click(object sender, EventArgs e)
		{
			
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//変更するデータを選択してください。
				MessageBox.Show("Pilih data yang ingin Anda ubah.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox8.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox4.Text))
			{
				//Basic Salaryが入力されていません。
				MessageBox.Show("Basic Salary belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox5.Text))
			{
				//Allowanceが入力されていません。
				MessageBox.Show("Allowance belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox6.Text))
			{
				//Transportが入力されていません。
				MessageBox.Show("Transport belum dimasukkan.");
				return;
			}

			using (var con = new NpgsqlConnection(connectionString))
			{
				DateTime today = DateTime.Today;

				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan " + textBox3.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "insert into basic_salary(e_code, b_salary, b_allowance, b_transport, b_startdate, b_update) " +
							"values(@code, @salary, @allowance, @transport, @startdate, @update);";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox2.Text));
							cmd.Parameters.AddWithValue("@salary", int.Parse(textBox4.Text));
							cmd.Parameters.AddWithValue("@allowance", int.Parse(textBox5.Text));
							cmd.Parameters.AddWithValue("@Transport", int.Parse(textBox6.Text));
							cmd.Parameters.AddWithValue("@startdate", DateTime.Parse(textBox8.Text));
							cmd.Parameters.AddWithValue("@update", today);
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
		private void reflesh()
		{
			

		}
		private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
		{
			textBox8.Text = e.Start.ToString("dd/MM/yyyy");
		}

		List<BS>bsList= new List<BS>();
		private void button2_Click(object sender, EventArgs e)
		{
			//同じ従業員番号のデータの選別をおこなう、将来の予約とまだ適用中
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			try
			{
				con.Open();
				// データ取得クエリ
				string sql = "SELECT * FROM basic_salary where b_flag is null;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					bsList.Add(new BS(
						int.Parse(reader["d_code"].ToString()),
						reader["e_code"].ToString(),
						long.Parse(reader["b_salary"].ToString()),
						long.Parse(reader["b_allowance"].ToString()),
						long.Parse(reader["b_transport"].ToString()),
						long.Parse(reader["b_startdate"].ToString()),
						reader["b_update"].ToString()));
				}
				reader.Close();
			}
			catch (Exception ex)
			{
				//問題が発生しリストを取得できませんでした。
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan." + ex.Message);
			}
			finally
			{
				con.Close();
			}

			// e_code でグループ化して同じ e_code を持つデータを特定
			var groupedData = bsList
				.GroupBy(bs => bs.E_code)
				.Where(group => group.Count() > 1); // 同じ e_code が複数あるグループのみ
			DateTime today = DateTime.Now;
			foreach (var group in groupedData)
			{
				// 今日の日付よりも B_update が以前のデータをフィルタ
				var filteredData = group
					.Where(bs => DateTime.Parse(bs.B_update) < today)
					.OrderByDescending(bs => DateTime.Parse(bs.B_update)) // 日付順で降順
					.ToList();

				// 一番新しいデータ以外を更新
				foreach (var bs in filteredData.Skip(1)) // 2番目以降
				{
					string updateSql = "UPDATE basic_salary SET b_flag = 1 WHERE b_code = @b_code;";
					using (NpgsqlCommand updateCom = new NpgsqlCommand(updateSql, con))
					{
						updateCom.Parameters.AddWithValue("@b_code", bs.B_code);
						updateCom.ExecuteNonQuery();
					}
				}
			}
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
						string sql = "INSERT INTO basic_salary(e_code, b_salary, b_allowance, b_transport,b_startdate, b_update)" +
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

		private void textBox2_Leave(object sender, EventArgs e)
		{

		}
	}
}
