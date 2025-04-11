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
	public partial class Cal_01 : Userbase
	{
		public Cal_01()
		{
			InitializeComponent();
		}
		List<P_Info> piList = new List<P_Info>();
		private void label5_Click(object sender, EventArgs e)
		{

		}

		private void Cal_01_Load(object sender, EventArgs e)
		{
			//Personal_Infoの情報をリストに入れておく
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
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
			//combobox1にタイプを入れる。0:Rapelan, 1:adj, 2:other
			comboBox1.Items.Add("Rapelan");
			comboBox1.Items.Add("Adj");
			comboBox1.Items.Add("Other");

			renewal_sh();
		}
		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code");
			dt1.Columns.Add("Code Karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Tipe");
			dt1.Columns.Add("Amount");
			dt1.Columns.Add("Remark");
			dt1.Columns.Add("Date");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
			try
			{
				con.Open();

				// データ取得クエリ
				string sql = "SELECT DISTINCT other.o_code, other.e_code, personal_info.p_name, other.o_type, " +
					"other.o_amount, other.o_remark, other.o_paydate FROM other left join personal_info on " +
					"other.e_code =personal_info.p_code";
				
				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["o_code"],
						reader["e_code"],
						reader["p_name"],
						reader["o_type"],
						reader["o_amount"],
						reader["o_remark"],
						reader["o_paydate"]
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

		private void label2_Click(object sender, EventArgs e)
		{

		}

		private void textBox3_Leave(object sender, EventArgs e)
		{
			try
			{
				int inputCode;
				if (int.TryParse(textBox3.Text, out inputCode))
				{
					// piListから一致するPI_codeを検索
					var matchedItem = piList.FirstOrDefault(p => p.PI_code == inputCode);

					if (matchedItem != null)
					{
						// textBox3にPI_nameを表示
						textBox4.Text = matchedItem.PI_name;
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

		private void button1_Click(object sender, EventArgs e)
		{
			//空欄判定
			if (string.IsNullOrWhiteSpace(textBox3.Text))
			{
				//Code Karyawanが選択されていません。
				MessageBox.Show("Code Karyawan tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(comboBox1.Text))
			{
				//Tipeが選択されていません。
				MessageBox.Show("Tipe tidak dipilih.");
				return;
			}
			if (string.IsNullOrWhiteSpace(textBox5.Text))
			{
				//合計金額が記入されていません。
				MessageBox.Show("Jumlah total tidak diisi.");
				return;
			}
			using (var con = new NpgsqlConnection(connectionString))
			{
				DialogResult kekka;
				//textBox2の内容を登録しますか？
				kekka = MessageBox.Show(this, "Apakah Anda ingin mendaftarkan " + textBox4.Text + "?", "*perhatian*", MessageBoxButtons.YesNo);
				if (kekka == DialogResult.Yes)
				{
					con.Open();
					try
					{
						string query = "insert into other(e_code, o_type, o_amount, o_remark, o_paydate) values(@code, @type, @amount, @remark, @paydate)";

						using (var cmd = new NpgsqlCommand(query, con))
						{
							// パラメータを設定
							cmd.Parameters.AddWithValue("@code", int.Parse(textBox3.Text));
							cmd.Parameters.AddWithValue("@type", comboBox1.SelectedIndex);
							cmd.Parameters.AddWithValue("@amount", int.Parse(textBox5.Text));
							cmd.Parameters.AddWithValue("@remark", textBox2.Text);
							cmd.Parameters.AddWithValue("@paydate", dateTimePicker1.Value);

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
			if (string.IsNullOrWhiteSpace(textBox1.Text))
			{
				//Codeが選択されていません。
				MessageBox.Show("Code tidak dipilih.");
				return;
			}
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
							string query = "update other set e_code=@code, o_amount=@amount, o_remark=@remark" +
								"o_paydate=@paydate where o_code=@o";

							using (var cmd = new NpgsqlCommand(query, con))
							{
								// パラメータを設定
								cmd.Parameters.AddWithValue("@o", int.Parse(textBox1.Text));
								cmd.Parameters.AddWithValue("@code", int.Parse(textBox3.Text));
								cmd.Parameters.AddWithValue("@amount", int.Parse(textBox5.Text));
								cmd.Parameters.AddWithValue("@remark", textBox2.Text);
								cmd.Parameters.AddWithValue("@paydate", dateTimePicker1.Value);

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

		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			try
			{
				textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
				textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
				textBox4.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
				comboBox1.SelectedIndex=int.Parse(dataGridView1.CurrentRow.Cells[3].Value.ToString());
				textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
				dateTimePicker1.Value = DateTime.Parse(dataGridView1.CurrentRow.Cells[6].Value.ToString());
				textBox2.Text= dataGridView1.CurrentRow.Cells[5].Value.ToString();
			}
			catch { }
		}

		private void button3_Click(object sender, EventArgs e)
		{

		}


	}
}
