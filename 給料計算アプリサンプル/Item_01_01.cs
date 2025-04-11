using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;

namespace 給料計算アプリサンプル
{
	public partial class Item_01_01 : Userbase
	{
		
		
		public Item_01_01()
		{
			InitializeComponent();
		}

		private void label9_Click(object sender, EventArgs e)
		{

		}

		private void button4_Click(object sender, EventArgs e)
		{
			Item_01_02 fm2 = new Item_01_02();
			fm2.Show();
		}
		List<Dept> deptList = new List<Dept>();  //combobox1
		List<Dept> deptList2 = new List<Dept>();
		List<Title> titleList = new List<Title>();
		List<Title> titleList2 = new List<Title>();

		private void Item_01_01_Load(object sender, EventArgs e)
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

				comboBox3.Items.Add("Bekerja");
				comboBox3.Items.Add("Pensiun");
				comboBox3.SelectedIndex = 0;

				comboBox6.Items.Add("Bekerja");
				comboBox6.Items.Add("Pensiun");


				comboBox7.Items.Add("Peserta");
				comboBox7.Items.Add("Pensiun");
				renewal_sh();
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

		DataTable dt1;
		private void renewal_sh()
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
			dt1.Columns.Add("Bank");
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
					sql = sql + "p_code=like %" + textBox1.Text + "%";
				}
				if (!string.IsNullOrWhiteSpace(textBox4.Text))
				{

					if (i == 0)
					{
						sql = sql + " where p_name=like %" + textBox4.Text + "% ";
					}
					else
					{
						sql = sql + " and p_name=like %" + textBox4.Text + "% ";
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
						sql = sql + " and p_title=" + comboBox2.SelectedValue.ToString();
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
						reader["p_bank"],
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
			try
			{
				flag();
			}
			catch { }
		}

		//p_flagにフラグを立てる
		private void flag()
		{
			using (var connection = new NpgsqlConnection(connectionString))
			{
				connection.Open();

				// 今日の日付を取得
				DateTime today = DateTime.Today;

				// 最も古い p_startdate を取得
				string getOldestDateQuery = @"
                SELECT p_startdate
                FROM public.personal_info
                WHERE p_startdate < @today
                ORDER BY p_startdate ASC
                LIMIT 1;";

				DateTime? oldestDate = null;

				using (var command = new NpgsqlCommand(getOldestDateQuery, connection))
				{
					command.Parameters.AddWithValue("@today", today);

					using (var reader = command.ExecuteReader())
					{
						if (reader.Read())
						{
							oldestDate = reader.GetDateTime(0);
						}
					}
				}

				if (oldestDate.HasValue)
				{
					// 最も古いデータ以外の p_flag を更新
					string updateQuery = @"
                    UPDATE public.personal_info
                    SET p_flag = 1
                    WHERE p_startdate < @today AND p_startdate <> @oldestDate;
                ";

					using (var updateCommand = new NpgsqlCommand(updateQuery, connection))
					{
						updateCommand.Parameters.AddWithValue("@today", today);
						updateCommand.Parameters.AddWithValue("@oldestDate", oldestDate.Value);

						int rowsAffected = updateCommand.ExecuteNonQuery();
						Console.WriteLine($"{rowsAffected} rows updated.");
					}
				}
				else
				{
					Console.WriteLine("No records found with p_startdate earlier than today.");
				}
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox3.Text))
			{
				//Basic Salaryが入力されていません。
				MessageBox.Show("Name belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox4.Text))
			{
				//Deptが選択されていません。
				MessageBox.Show("Dept tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox5.Text))
			{
				//Titleが選択されていません。
				MessageBox.Show("Title tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(dateTimePicker1.Text))
			{
				//日付が選択されていません。
				MessageBox.Show("Tanggal mulai yang berlaku tidak dipilih.");
				return;
			}

			if (string.IsNullOrWhiteSpace(comboBox6.Text))
			{
				//Statusが選択されていません。
				MessageBox.Show("Status tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox7.Text))
			{
				//SPが選択されていません。
				MessageBox.Show("SP tidak dipilih.");
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
						string query = "insert into personal_info(p_code, p_name, p_dept, p_title, p_mail, p_status, p_sp, p_bank, p_update, p_startdate) " +
							"values(@code, @name, @dept, @title, @mail, @status, @sp, @bank, @update, @start);";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox2.Text));
							cmd.Parameters.AddWithValue("@name", textBox3.Text);
							cmd.Parameters.AddWithValue("@dept", comboBox4.SelectedValue);
							cmd.Parameters.AddWithValue("@title", int.Parse(comboBox5.SelectedValue.ToString()));
							cmd.Parameters.AddWithValue("@mail", textBox5.Text);
							cmd.Parameters.AddWithValue("@status", comboBox6.SelectedIndex);
							cmd.Parameters.AddWithValue("@sp", comboBox7.SelectedIndex);
							cmd.Parameters.AddWithValue("@bank", textBox6.Text);
							cmd.Parameters.AddWithValue("@update", today);
							cmd.Parameters.AddWithValue("@start", dateTimePicker1.Value);
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

		private void button1_Click(object sender, EventArgs e)
		{
			renewal_sh();
		}

		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{

			
			try
			{
				// 各セルが null または空文字の場合は空白をセット
				textBox2.Text = dataGridView1.CurrentRow.Cells[0].Value?.ToString().Trim() == "" ? "" : dataGridView1.CurrentRow.Cells[0].Value.ToString(); // 従業員コード
				textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value?.ToString().Trim() == "" ? "" : dataGridView1.CurrentRow.Cells[1].Value.ToString(); // 従業員名
        
				comboBox4.SelectedValue = (dataGridView1.CurrentRow.Cells[2].Value != null && int.TryParse(dataGridView1.CurrentRow.Cells[2].Value.ToString(), out int dep))
					? dep : -1; // 部門選択

				comboBox5.SelectedValue = (dataGridView1.CurrentRow.Cells[3].Value != null && int.TryParse(dataGridView1.CurrentRow.Cells[3].Value.ToString(), out int pos))
					? pos : -1; // 肩書選択

				textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value?.ToString().Trim() == "" ? "" : dataGridView1.CurrentRow.Cells[4].Value.ToString(); // メールアドレス
        
				comboBox6.SelectedIndex = (dataGridView1.CurrentRow.Cells[5].Value != null && int.TryParse(dataGridView1.CurrentRow.Cells[5].Value.ToString(), out int work))
					? work : -1; // 就労状況

				comboBox7.SelectedIndex = (dataGridView1.CurrentRow.Cells[6].Value != null && int.TryParse(dataGridView1.CurrentRow.Cells[6].Value.ToString(), out int union))
					? union : -1; // 組合加入状況

				textBox6.Text = dataGridView1.CurrentRow.Cells[7].Value?.ToString().Trim() == "" ? "" : dataGridView1.CurrentRow.Cells[7].Value.ToString(); // 銀行口座
        
				dateTimePicker1.Value = (dataGridView1.CurrentRow.Cells[8].Value != null && DateTime.TryParse(dataGridView1.CurrentRow.Cells[8].Value.ToString(), out DateTime date))
					? date : DateTime.Today; // 適用開始日
			}
			catch { }
			
		}


	}
}
