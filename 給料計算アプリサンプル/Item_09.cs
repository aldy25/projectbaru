using Npgsql;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace 給料計算アプリサンプル
{
	public partial class Item_09 : Userbase
	{
		public Item_09()
		{
			InitializeComponent();
		}

		private void Item_09_Load(object sender, EventArgs e)
		{
			renewal_sh();

			comboBox1.Items.Add("katering");
			comboBox1.Items.Add("Admin");
		}

		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code");
			dt1.Columns.Add("Name Kontak");
			dt1.Columns.Add("WhatsApp");
			dt1.Columns.Add("Tipe");
			dt1.Columns.Add("Tanggal dimasukkan");


			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT * FROM contact";
				
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["ct_code"],
						reader["ct_name"],
						reader["ct_whatsapp"],
						reader["ct_type"],
						reader["ct_entrydate"]
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
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox1.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox2.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox3.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox1.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
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
						string query = "update contact set ct_name=@name, ct_whatsapp=@whatsapp, ct_type=@type, ct_entrydate=@entrydate " +
							"where ct_code=@code;";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox1.Text));
							cmd.Parameters.AddWithValue("@name", textBox2.Text);
							cmd.Parameters.AddWithValue("@whatsapp", textBox3.Text);
							cmd.Parameters.AddWithValue("@type", comboBox1.SelectedIndex);
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
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox1.Text))
			{
				//Codeが入力されていません。
				MessageBox.Show("Date belum dimasukkan.");
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
						string query = "insert into contact(ct_name,ct_whatsapp,ct_type,ct_entrydate) values(@name, @whatsapp, @type, @entrydate);";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@name", textBox2.Text);
							cmd.Parameters.AddWithValue("@whatsapp", textBox3.Text);
							cmd.Parameters.AddWithValue("@type", comboBox1.SelectedIndex);
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
				}
			}
		}

		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			try
			{
				textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
				textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
				textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
				comboBox1.SelectedIndex = int.Parse(dataGridView1.CurrentRow.Cells[3].Value.ToString());
			}
			catch { }
		}
	}
}
