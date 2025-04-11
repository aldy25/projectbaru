using iText.StyledXmlParser.Jsoup.Select;
using MathNet.Numerics;
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
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static NPOI.HSSF.Util.HSSFColor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace 給料計算アプリサンプル
{

	public partial class Cal_02 : Userbase
	{
		public Cal_02()
		{
			InitializeComponent();
		}
		List<Dept> deptList = new List<Dept>();  //combobox1
		List<Title> titleList = new List<Title>();//combobox2
		List<Payroll> payList = new List<Payroll>();//給料明細データ持ち越し用

		private void Cal_02_Load(object sender, EventArgs e)
		{
			//Combobox1にDeptの情報を入れる。
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
			catch
			{

			}

			// 現在の日付を取得
			DateTime now = DateTime.Now;

			// datetimepicker1: 先月の21日
			dateTimePicker1.Value = new DateTime(now.Year, now.Month, 1).AddMonths(-1).AddDays(20);

			// datetimepicker2: 今月の20日
			dateTimePicker2.Value = new DateTime(now.Year, now.Month, 20);

			// datetimepicker3: 今月の月末
			dateTimePicker3.Value = new DateTime(now.Year, now.Month, DateTime.DaysInMonth(now.Year, now.Month));

			renewal_sh();
		}

		public int workingDays = 0;
		private void CalculateWorkingDays()
		{
			// DateTimePickerから日付を取得
			DateTime startDate = dateTimePicker1.Value.Date;
			DateTime endDate = dateTimePicker2.Value.Date;

			int workingDays = 0;

			using (var conn = new NpgsqlConnection(connectionString))
			{
				conn.Open();
				string query = @"
                SELECT wd_date FROM public.working_days
                WHERE wd_tipe = 1 AND wd_date BETWEEN @startDate AND @endDate";

				using (var cmd = new NpgsqlCommand(query, conn))
				{
					cmd.Parameters.AddWithValue("@startDate", startDate);
					cmd.Parameters.AddWithValue("@endDate", endDate);

					using (var reader = cmd.ExecuteReader())
					{
						// データベースから取得した稼働日リスト
						var dbWorkingDays = new HashSet<DateTime>();

						while (reader.Read())
						{
							dbWorkingDays.Add(reader.GetDateTime(0));
						}

						// 指定期間の日付をチェック
						for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
						{
							if (dbWorkingDays.Contains(date))
							{
								// すでにDBに登録されている日付
								workingDays++;
							}
							else
							{
								// データがない場合、土日かどうか判定
								if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
								{
									workingDays++; // 平日の場合、労働日としてカウント
								}
							}
						}
					}
				}
			}
			// 結果を表示
			textBox3.Text = workingDays.ToString();
		}
		private void button1_Click(object sender, EventArgs e)
		{
			renewal_sh();
		}
		DataTable dt1;
		private void renewal_sh()
		{
			// dataGridView1に部門の情報を表示する
			dt1 = new DataTable();
			dt1.Columns.Add("Code Karyawan");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Jumlah OT");
			dt1.Columns.Add("Hari kerja");
			dt1.Columns.Add("Cuti");
			dt1.Columns.Add("Cuti Sakit");
			dt1.Columns.Add("liburan yang tidak dibayar");

			// PostgreSQLの接続文字列
			NpgsqlConnection con = new NpgsqlConnection(connectionString);
		
			try
			{
				con.Open();
				int whereQty=0;//where句があるかどうか判定
				// データ取得クエリ
				string StartDate = dateTimePicker1.Value.ToString("yyyy-MM-dd");//給料計算開始日
				string FinishDate = dateTimePicker2.Value.ToString("yyyy-MM-dd");//給料計算終了日

				string sql =
					"WITH date_range AS(SELECT generate_series('"+ StartDate + "'::date, '"+ FinishDate + "'::date, '1 day'::interval)::date AS date)," +
					" workdays AS(SELECT DISTINCT ws.w_date FROM public.working_status ws WHERE ws.w_date BETWEEN '"+ StartDate + "' AND '"+ FinishDate + "')," +
					" overtime_summary AS(SELECT ws.e_code,SUM(ws.w_overtime) AS total_overtime FROM working_status ws" +
					" WHERE ws.w_date BETWEEN '"+ StartDate + "' AND '"+ FinishDate + "' GROUP BY ws.e_code)," +
					" day_off_summary AS(SELECT dof.e_code,SUM(CASE WHEN dof.d_tipe = 0 THEN 1 ELSE 0 END) AS tipe_0_count," +
					" SUM(CASE WHEN dof.d_tipe = 1 THEN 1 ELSE 0 END) AS tipe_1_count,SUM(CASE WHEN dof.d_tipe = 2 THEN 1 ELSE 0 END) AS tipe_2_count" +
					" FROM public.day_off dof WHERE dof.d_date BETWEEN '"+StartDate+"' AND '"+ FinishDate + "' GROUP BY dof.e_code)" +
					" SELECT DISTINCT pi.p_code, pi.p_name, COALESCE(os.total_overtime, 0) AS total_overtime, (SELECT COUNT(*) FROM workdays) AS workdays_count," +
					" COALESCE(dos.tipe_0_count, 0) AS tipe_0_count, COALESCE(dos.tipe_1_count, 0) AS tipe_1_count, COALESCE(dos.tipe_2_count, 0) AS tipe_2_count" +
					" FROM public.personal_info pi LEFT JOIN overtime_summary os ON pi.p_code = os.e_code LEFT JOIN day_off_summary dos ON pi.p_code = dos.e_code";
				if (!string.IsNullOrWhiteSpace(textBox1.Text))
				{
					if (whereQty > 0)
					{
						sql += " AND pi.p_code LIKE '%" + textBox1.Text + "%'";

					}
					else
					{
						sql += " WHERE pi.p_code LIKE '%" + textBox1.Text + "%'";
					}
					whereQty++;
				}

				if (!string.IsNullOrWhiteSpace(textBox2.Text))
				{
					if (whereQty > 0)
					{
						sql += " AND pi.p_name LIKE '%" + textBox2.Text + "%'";
					}
					else
					{
						sql += " WHERE pi.p_name LIKE '%" + textBox2.Text + "%'";
					}
					whereQty++;
				}

				if (comboBox1.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox1.SelectedValue.ToString()))
				{
					if (whereQty > 0)
					{
						sql += " AND pi.p_dept = '" + comboBox1.SelectedValue.ToString() + "'";
					}
					else
					{
						sql += " WHERE pi.p_dept = '" + comboBox1.SelectedValue.ToString() + "'";
					}
					whereQty++;
				}

				if (comboBox2.SelectedValue != null && !string.IsNullOrWhiteSpace(comboBox2.SelectedValue.ToString()))
				{
					if (whereQty > 0)
					{
						sql += " AND pi.p_title = '" + comboBox2.SelectedValue.ToString() + "'";
					}
					else
					{
						sql += " WHERE pi.p_title = '" + comboBox2.SelectedValue.ToString() + "'";
					}
					whereQty++;
				}

				NpgsqlCommand com = new NpgsqlCommand(sql, con);
				NpgsqlDataReader reader = com.ExecuteReader();
				while (reader.Read())
				{
					dt1.Rows.Add(
						reader["p_code"],
						reader["p_name"],
						reader["total_overtime"],
						reader["workdays_count"],
						reader["tipe_0_count"],
						reader["tipe_1_count"],
						reader["tipe_2_count"]
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
			CalculateWorkingDays();
			
		}

		private void button2_Click(object sender, EventArgs e)
		{
			//基本データ取得
			try
			{
				NpgsqlConnection con = new NpgsqlConnection(connectionString);
				for (int i = 0; i < dt1.Rows.Count; i++)
				{
					int p_code =0;          //従業員コード
					string Name = "";       //従業員名前
					string Dept = "";       //部門名
					string Title = "";		//肩書
					string Bank = "";		//銀行口座
					long Basic = 0;			//基本給
					long b_allowance = 0;	//手当
					long b_transport = 0;	//交通費
					long o_rapelan =0;      //Rapelanの金額
					long o_adj = 0;         //Adjの金額
					long o_others = 0;		//Otherの金額
					long absence = 0;		//欠勤費用の差引,specialLeave
					long Other=0;			//その他費用
					double BPJSkeshR = 0;
					double BPJSNakarR = 0;
					double BPJSpenR = 0;
					long BPJSkesh = 0;
					long BPJSNakar = 0;
					long BPJSpen = 0;
					long SPcost = 0;		//組合費
					double OTA = 0;			//残業代1.5倍
					double OTB = 0;         //残業代2.0倍
					double OTC = 0;         //残業代3.0倍
					double OTD = 0;         //残業代4.0倍

					try
					{
						p_code=int.Parse(dt1.Rows[i][0].ToString());
						try
						{
							using (var conn = new NpgsqlConnection(connectionString))
							{
								conn.Open();
								string sql = "SELECT p.p_name, p.p_bank, t.t_name FROM personal_info p left join title t " +
									"ON p.p_title=t.t_code WHERE p.p_code = @p_code";

								using (var cmd = new NpgsqlCommand(sql, conn))
								{
									cmd.Parameters.AddWithValue("@p_code", p_code);

									using (var reader = cmd.ExecuteReader())
									{
										if (reader.Read()) // データがある場合
										{
											Name = reader.GetString(0);
											Title= reader.GetString(2);
											Bank = reader.GetString(1);
										}
									}
								}
							}
						}
						catch { }
						//部門名取得
						try
						{
							using (var conn = new NpgsqlConnection(connectionString))
							{
								conn.Open();
								string sql = "SELECT p.p_code, d.dp_name AS Dept FROM personal_info p LEFT JOIN department d ON p.p_dept = d.dp_code;";

								using (var cmd = new NpgsqlCommand(sql, conn))
								{
									cmd.Parameters.AddWithValue("@p_code", p_code);

									using (var reader = cmd.ExecuteReader())
									{
										if (reader.Read()) // データがある場合
										{
											Dept = reader.GetString(1);
										}
									}
								}
							}
						}
						catch { }
						
						//基本給・手当のデータを取り出す
						try
						{
							con.Open();
							string sql = "SELECT b_salary,b_allowance, b_transport FROM basic_salary  where e_code=" + p_code + " AND b_startdate <= CURRENT_DATE ORDER BY b_startdate DESC LIMIT 1;;";
							NpgsqlCommand com = new NpgsqlCommand(sql, con);
							NpgsqlDataReader reader = com.ExecuteReader();
							while (reader.Read())
							{
								Basic = reader.GetInt64(0);
								b_allowance= reader.GetInt64(1);
								b_transport = reader.GetInt64(2);
							}
						}
						catch
						{

						}
						finally
						{
							con.Close();
						}
						//交通費の金額を取得
						try
						{
							con.Open();
							string sql = "SELECT c_transport FROM company where c_update <= CURRENT_DATE ORDER BY c_update DESC LIMIT 1;;";
							NpgsqlCommand com = new NpgsqlCommand(sql, con);
							NpgsqlDataReader reader = com.ExecuteReader();
							while (reader.Read())
							{
								b_transport = reader.GetInt64(0);
							}
						}
						catch
						{

						}
						finally
						{
							con.Close();
						}
						//その他の金額を取得する。
						try
						{
							using (var conn = new NpgsqlConnection(connectionString))
							{
								conn.Open();
								string sql = "SELECT o_type, o_amount FROM other WHERE e_code = @p_code and o_paydate=@paydate";

								using (var cmd = new NpgsqlCommand(sql, conn))
								{
									cmd.Parameters.AddWithValue("@p_code", p_code);
									cmd.Parameters.AddWithValue("@paydate", dateTimePicker3.Value.ToString("dd MMMM yyyy"));
									using (var reader = cmd.ExecuteReader())
									{
										if (reader.Read()) // データがある場合
										{
											switch (reader.GetInt32(0))
											{
												case 0:
													o_rapelan = reader.GetInt32(1);
													break;
												case 1:
													o_adj = reader.GetInt32(1);
													break;
												case 2:
													o_others= reader.GetInt32(1);
													break;
											}
										}
									}
								}
							}
						}
						catch { }

						//残業代の計算OT
						try
						{
							using (var conn = new NpgsqlConnection(connectionString))
							{
								conn.Open();

								// 指定した期間のデータを取得
								string sql = @"
								SELECT w.w_date, w.w_overtime, 
								(CASE WHEN ph.ph_date IS NOT NULL OR EXTRACT(DOW FROM w.w_date) IN (0,6) 
								THEN 'holiday' ELSE 'weekday' END) AS day_type
								FROM public.working_status w
								LEFT JOIN public.public_holiday ph ON w.w_date = ph.ph_date
									WHERE w.w_date BETWEEN @startDate AND @endDate AND w.e_code = @ecode";

								using (var cmd = new NpgsqlCommand(sql, conn))
								{
									cmd.Parameters.AddWithValue("startDate", dateTimePicker1.Value);
									cmd.Parameters.AddWithValue("endDate", dateTimePicker2.Value);
									cmd.Parameters.AddWithValue("ecode", p_code);

									using (var reader = cmd.ExecuteReader())
									{
										while (reader.Read())
										{
											double overtime = reader.IsDBNull(1) ? 0 : reader.GetDouble(1); // 分単位
											string dayType = reader.GetString(2);
											double otHours = overtime / 60.0; // 分→時間に変換

											if (dayType == "weekday")
											{
												OTA += Math.Min(1, otHours);
												OTB += Math.Min(7, Math.Max(0, otHours - 1));
												OTC += Math.Min(1, Math.Max(0, otHours - 8));
												OTD += Math.Max(0, otHours - 9);
											}
											else // 土日祝日
											{
												OTB += Math.Min(8, otHours);
												OTC += Math.Min(1, Math.Max(0, otHours - 8));
												OTD += Math.Max(0, otHours - 9);
											}
										}

									}
								}
							}
						}
						catch { }

						//交通費の計算
						try
						{
							b_transport = b_transport * (int.Parse(dt1.Rows[i][3].ToString())- int.Parse(dt1.Rows[i][4].ToString())- int.Parse(dt1.Rows[i][5].ToString()) - int.Parse(dt1.Rows[i][4].ToString()) - int.Parse(dt1.Rows[i][6].ToString()));
						}
						catch { }

						//欠勤費用の差引
						try
						{
							absence = Basic/ int.Parse(textBox3.Text) * int.Parse(dt1.Rows[i][6].ToString());
						}
						catch { }
						//その他費用
						try
						{
							DateTime selectedDate = dateTimePicker3.Value.Date;
							using (var conn = new NpgsqlConnection(connectionString))
							{
								conn.Open();
								string sql = "SELECT COALESCE(SUM(o_amount), 0) FROM other WHERE e_code = @p_code AND o_paydate = @selectedDate";

								using (var cmd = new NpgsqlCommand(sql, conn))
								{
									cmd.Parameters.AddWithValue("@p_code", p_code);
									cmd.Parameters.AddWithValue("@selectedDate", selectedDate);

									var result = cmd.ExecuteScalar();
									if (result != DBNull.Value)
									{
										Other = Convert.ToInt64(result);
									}
								}
							}
						}
						catch { }
						//BPJS
						using (var conn = new NpgsqlConnection(connectionString))
						{
							conn.Open();

							// 指定した期間のデータを取得
							string sql = @"Select bp_kesh, bp_naker, bp_pensiun from bpjs where bp_update <= CURRENT_DATE;";
							using (var cmd = new NpgsqlCommand(sql, conn))
							using (var reader = cmd.ExecuteReader())
							{
								if (reader.Read())
								{
									BPJSkeshR = reader.GetDouble(0);
									BPJSkesh= (long)Math.Floor((Basic+b_allowance) * BPJSkeshR/100);

									BPJSNakarR = reader.GetDouble(1);
									BPJSNakar = (long)Math.Floor((Basic + b_allowance) * BPJSNakarR/100);

									BPJSpenR = reader.GetDouble(2);
									BPJSpen = (long)Math.Floor((Basic + b_allowance) * BPJSpenR / 100);
								}
							}

						}

						//組合費
						using (var connection = new NpgsqlConnection(connectionString))
						{
							connection.Open();

							// SQL クエリ
							string sql = @"
								WITH latest_member AS(
									SELECT p_code, p_name, p_sp, p_startdate
									FROM public.personal_info
									WHERE p_code = @p_code
									AND p_sp = 1
									AND p_startdate<CURRENT_DATE
									ORDER BY p_startdate DESC
									LIMIT 1)
								SELECT
									c.c_spcost
								FROM
									latest_member lm
								CROSS JOIN
									public.company c
								LIMIT 1;";

							using (var command = new NpgsqlCommand(sql, connection))
							{
								// パラメータを追加
								command.Parameters.AddWithValue("@p_code", p_code);

								using (var reader = command.ExecuteReader())
								{
									if (reader.Read()) // 結果があれば取得
									{
										SPcost = reader.GetInt64(0);
									}
								}
							}
						}
						//payListにデータを入力
						payList.Add(new Payroll(p_code,  Name, Dept, Title, Bank, Basic, b_allowance,b_transport, o_rapelan, o_adj, absence, Other, BPJSkesh, BPJSNakar, BPJSpen, SPcost, OTA, OTB, OTC, OTD));
					}
					catch
					{

					}
				}
			}
			catch
			{
				//基本データの取得に失敗しました。開発者に連絡してください。
				MessageBox.Show("Gagal mengambil data dasar.Tolong hubungi pengembang.");
				return;
			}

			//資料の作成
			//①給料振込データの作成

			try
			{
				List<ReBase.Payroll> convertedPayList = payList.Select(p => new ReBase.Payroll(p.P_code, p.Name, p.Dept,p.Title, p.Bank, p.Basic, p.Allowance, p.Transport,
									p.Absence, p.Other, p.BPJSkesh, p.BPJSNakar, p.BPJSpen, p.SPcost,p.OTA, p.OTB, p.OTC, p.OTD)).ToList();
				Re_01C re1 = new Re_01C(convertedPayList,workingDays);
				re1.makeReport1();
			}
			catch 
			{
				MessageBox.Show("Gagal membuat data transfer gaji.");
			}

			//②JACデータの作成
			try
			{
				List<ReBase.Payroll> convertedPayList = payList.Select(p => new ReBase.Payroll(p.P_code, p.Name, p.Dept, p.Title ,p.Bank, p.Basic, p.Allowance, p.Transport,
									p.Absence, p.Other, p.BPJSkesh, p.BPJSNakar, p.BPJSpen, p.SPcost, p.OTA, p.OTB, p.OTC, p.OTD)).ToList();
				Re_02C re2 = new Re_02C(convertedPayList, workingDays);
				re2.makeReport2();
			}
			catch 
			{
				MessageBox.Show("Gagal membuat data untuk JAC.");
			}

			//③給料明細データの作成、送信
			try
			{
				List<ReBase.Payroll> convertedPayList = payList.Select(p => new ReBase.Payroll(p.P_code, p.Name, p.Dept,p.Title, p.Bank, p.Basic, p.Allowance, p.Transport,
									p.Absence, p.Other, p.BPJSkesh, p.BPJSNakar, p.BPJSpen, p.SPcost, p.OTA, p.OTB, p.OTC, p.OTD)).ToList();
				Re_03C re3 = new Re_03C(convertedPayList,dateTimePicker1.Value, dateTimePicker2.Value, dateTimePicker3.Value, workingDays);
				re3.makeReport3();
			}
			catch 
			{
				MessageBox.Show("Gagal menghasilkan data Slip.");
			}
		}
	}
}
