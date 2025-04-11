using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 給料計算アプリサンプル
{
	public partial class Item_07 : Userbase
	{
		public Item_07()
		{
			InitializeComponent();
		}

		private void Item_07_Load(object sender, EventArgs e)
		{
			renewal_sh();
		}
		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("code");
			dt1.Columns.Add("Title");
			dt1.Columns.Add("Allowance");

			// PostgreSQLの接続文字列
			string connectionString = "Host=localhost;Port=5432;Username=postgres;Password=maruhide1950;Database=payroll";
			NpgsqlConnection con = new NpgsqlConnection(connectionString);

			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM title;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["t_code"],
						reader["t_name"],
						reader["t_allowance"]
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
				MessageBox.Show("Terjadi masalah dan daftar tidak dapat ditampilkan.\n" + ex.Message);
			}
			finally
			{
				con.Close();
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(textBox2.Text))
			{

			}
			else
			{
				//肩書が記入されていません。
				MessageBox.Show("Title tidak diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox3.Text))
			{
				if (int.TryParse(textBox3.Text, out int result))
				{
					
				}
				else
				{
					MessageBox.Show("Bukan angka.");
					return ;
				}
			}
			else
			{
				//手当が記入されていません。
				MessageBox.Show("Allowance belum diisi.");
				return;
			}

			string connectionString = "Host=localhost;Port=5432;Username=postgres;Password=maruhide1950;Database=payroll";
			using (var con = new NpgsqlConnection(connectionString))
			{
				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan " + textBox2.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "insert into title(T_name,t_allowance) values(@name, @allowance)";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@name", textBox2.Text);
							cmd.Parameters.AddWithValue("@allowance", int.Parse(textBox3.Text));
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
			textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
			textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
			textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
		}

		private void button2_Click(object sender, EventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(textBox1.Text))
			{

			}
			else
			{
				//変更したいコードが選択されていません。
				MessageBox.Show("Kode yang ingin Anda ubah tidak dipilih.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox2.Text))
			{

			}
			else
			{
				//肩書が記入されていません。
				MessageBox.Show("Title belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox2.Text))
			{

			}
			else
			{
				//手当が記入されていません。
				MessageBox.Show("Allowane belum diisi.");
				return;
			}

			string connectionString = "Host=localhost;Port=5432;Username=postgres;Password=maruhide1950;Database=payroll";
			using (var con = new NpgsqlConnection(connectionString))
			{
				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan" + textBox2.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "update title set T_name=@name, T_Allowance=@allowance where T_code=@code;";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox1.Text));
							cmd.Parameters.AddWithValue("@name", textBox2.Text);
							cmd.Parameters.AddWithValue("@allowance", int.Parse(textBox3.Text));

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
	}
}
