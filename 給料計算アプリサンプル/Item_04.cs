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
	public partial class Item_04 : Userbase
	{
		public Item_04()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			if (!string.IsNullOrWhiteSpace(textBox1.Text))
			{

			}
			else
			{
				//食事手当が記入されていません。
				MessageBox.Show("Meal allowance belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox6.Text))
			{

			}
			else
			{
				//食事手当(ドライバー)が記入されていません。
				MessageBox.Show("Meal allowance Driver belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox2.Text))
			{

			}
			else
			{
				//有給日数が記入されていません。
				MessageBox.Show("Hari cuti belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox3.Text))
			{

			}
			else
			{
				//組合費が記入されていません。
				MessageBox.Show("Biaya SP belum diisi.");
				return;
			}
			if (!string.IsNullOrWhiteSpace(textBox4.Text))
			{

			}
			else
			{
				//適用開始日の名前が記入されていません。
				MessageBox.Show("Tanggal mulai yang berlaku belum diisi.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox5.Text))
			{
				//交通費が入力されていません。
				MessageBox.Show("Tidak ada biaya transportasi yang dimasukkan.");
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
						string query = "insert into company(c_mealallowance, c_mealallowanceD, paid_horiday, c_spcost, c_update, c_transport) values(@meal,@meald, @paid, @spcost,@update, @transport);";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@meal", int.Parse(textBox1.Text));
							cmd.Parameters.AddWithValue("@meald", int.Parse(textBox6.Text));
							cmd.Parameters.AddWithValue("@paid", int.Parse(textBox2.Text));
							cmd.Parameters.AddWithValue("@spcost", int.Parse(textBox3.Text));
							cmd.Parameters.AddWithValue("@update", DateTime.Parse(textBox4.Text));
							cmd.Parameters.AddWithValue("@transport", int.Parse(textBox5.Text));
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

		private void Item_04_Load(object sender, EventArgs e)
		{
			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			try
			{
				con.Open();
				string today = DateTime.Today.ToString("dd/MM/yyyy");
				// データ取得クエリ
				string sql = "SELECT * FROM company where c_update <='"+today+"' ORDER BY c_update DESC LIMIT 1;";
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					textBox1.Text = reader["c_mealallowance"].ToString();
					textBox6.Text = reader["c_mealallowanced"].ToString();
					textBox2.Text = reader["paid_horiday"].ToString();
					textBox3.Text = reader["c_spcost"].ToString();
					textBox4.Text = DateTime.Parse(reader["c_update"].ToString()).ToString("dd/MM/yyyy");
					textBox5.Text= reader["c_transport"].ToString();
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
	}
}
