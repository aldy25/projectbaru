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
	public partial class Item_03 : Userbase
	{
		public Item_03()
		{
			InitializeComponent();
		}

		private void label1_Click(object sender, EventArgs e)
		{

		}

		private void Item_03_Load(object sender, EventArgs e)
		{
			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);

			try
			{
				con.Open();
				string today = DateTime.Today.ToString("dd/MM/yyyy");
				// データ取得クエリ
				string sql = "SELECT * FROM bpjs where bp_update <'" + today + "' ORDER BY bp_update DESC LIMIT 1;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					textBox1.Text = reader["bp_kesh"].ToString();
					textBox2.Text = reader["bp_naker"].ToString();
					textBox3.Text = reader["bp_pensiun"].ToString();
					textBox4.Text = DateTime.Parse(reader["bp_update"].ToString()).ToString("dd/MM/yyyy");
				}
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

		private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
		{
			textBox4.Text = e.Start.ToString("dd/MM/yyyy");
		}

		private void button1_Click(object sender, EventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(textBox1.Text))
			{

			}
			else
			{
				//BPJS keshが記入されていません。
				MessageBox.Show("BPJS kesh belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox2.Text))
			{

			}
			else
			{
				//BPJS nakerが記入されていません。
				MessageBox.Show("BPJS naker belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox3.Text))
			{

			}
			else
			{
				//組合費が記入されていません。
				MessageBox.Show("BPJS pensiun belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox4.Text))
			{

			}
			else
			{
				//適用開始日が記入されていません。
				MessageBox.Show("Tanggal mulai yang berlaku belum diisi.");
				return;
			}

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
						string query = "insert into bpjs(bp_kesh, bp_naker, bp_pensiun,bp_update) values(@kesh, @naker, @pensiun,@update);";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@kesh", int.Parse(textBox1.Text));
							cmd.Parameters.AddWithValue("@naker", int.Parse(textBox2.Text));
							cmd.Parameters.AddWithValue("@pensiun", int.Parse(textBox3.Text));
							cmd.Parameters.AddWithValue("@update", DateTime.Parse(textBox4.Text));
							// クエリを実行
							cmd.ExecuteNonQuery();
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
