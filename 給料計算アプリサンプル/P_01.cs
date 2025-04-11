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
	public partial class P_01 : Form
	{
		public P_01()
		{
			InitializeComponent();
		}

		public class MN
		{
			public string Code { get; set; }
			public string Name { get; set; }

			public MN(string code, string name)
			{
				this.Code = code;
				this.Name = name;
			}
		}
		List<MN> mnList = new List<MN>();
		Dictionary<string, UserControl> dict;

		private void P_01_Load(object sender, EventArgs e)
		{
			mnList.Add(new MN("1", "Info Privadi"));
			mnList.Add(new MN("1", "Basic Salary"));
			mnList.Add(new MN("1", "BPJS"));
			mnList.Add(new MN("1", "Peraturan PT"));
			mnList.Add(new MN("1", "Libur National"));
			mnList.Add(new MN("1", "Dept"));
			mnList.Add(new MN("1", "Title"));
			mnList.Add(new MN("1", "Shift"));
			mnList.Add(new MN("1", "Contact"));
			mnList.Add(new MN("1", "Jam kerja"));

			mnList.Add(new MN("2", "Finger"));
			mnList.Add(new MN("2", "Info Cuti"));
			mnList.Add(new MN("2", "Lembur"));
			mnList.Add(new MN("2", "Attend"));

			mnList.Add(new MN("3", "Lain"));
			mnList.Add(new MN("3", "Payroll"));
			mnList.Add(new MN("3", "Bonus"));

			mnList.Add(new MN("4", "Transfer"));
			mnList.Add(new MN("4", "Akuntan"));
			mnList.Add(new MN("4", "Slip"));
			mnList.Add(new MN("4", "Lembur Tetup"));
			mnList.Add(new MN("4", "Rekam Lembur"));

			dict = new Dictionary<string, UserControl>();
			Item_01_01 item_01_01 = new Item_01_01();
			dict.Add("Info Privadi", item_01_01);
			Item_02 item_02 = new Item_02();
			dict.Add("Basic Salary", item_02);
			Item_03 item_03 = new Item_03();
			dict.Add("BPJS", item_03);
			Item_04 item_04 = new Item_04();
			dict.Add("Peraturan PT", item_04);
			Item_05 item_05 = new Item_05();
			dict.Add("Libur National", item_05);
			Item_06 item_06 = new Item_06();
			dict.Add("Dept", item_06);
			Item_07 item_07 = new Item_07();
			dict.Add("Title", item_07);
			Item_08 item_08 = new Item_08();
			dict.Add("Shift", item_08);
			Item_09 item_09 = new Item_09();
			dict.Add("Contact", item_09);
			Item_10 item_10 = new Item_10();
			dict.Add("Jam kerja", item_10);

			Att_01 att_01 = new Att_01();
			dict.Add("Finger", att_01);
			Att_02_01 att_02 = new Att_02_01();
			dict.Add("Info Cuti", att_02);
			Att_03 att_03 = new Att_03();
			dict.Add("Lembur", att_03);
			Att_04 att_04 = new Att_04();
			dict.Add("Attend", att_04);

			Cal_01 cal_01 = new Cal_01();
			dict.Add("Lain", cal_01);
			Cal_02 cal_02 = new Cal_02();
			dict.Add("Payroll", cal_02);
			Cal_03 cal_03 = new Cal_03();
			dict.Add("Bonus", cal_03);

			Re_01 re_01 = new Re_01();
			dict.Add("Transfer", re_01);
			Re_02 re_02 = new Re_02();
			dict.Add("Akuntan", re_01);
			Re_03 re_03 = new Re_03();
			dict.Add("Slip", re_03);
			Re_04 re_04 = new Re_04();
			dict.Add("Lembur Tetup", re_04);
			Re_05 re_05 = new Re_05();
			dict.Add("Rekam Lembur", re_05);

		}
		int Bqty = 0; //ボタンの個数を保存する。
		System.Windows.Forms.Button[] button = new System.Windows.Forms.Button[10];
		public void button_setting(int Search)
		{
			try
			{
				//ボタンを削除する。
				if (Bqty > 0)
				{
					for (int i = 0; i < Bqty; i++)
					{
						this.Controls.Remove(button[i]);
						button[i].Dispose();
					}
					Bqty = 0;
				}

				List<MN> cat = mnList.FindAll(x => int.Parse(x.Code.ToString()) == Search);
				for (int i = 0; i < cat.Count; i++)
				{
					button[i] = new System.Windows.Forms.Button();
					this.button[i].Font = new System.Drawing.Font("MS UI Gothic", 12, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
					this.button[i].Location = new System.Drawing.Point(28, 10 + i * 49);
					this.button[i].Size = new System.Drawing.Size(145, 43);
					this.button[i].BackColor = System.Drawing.SystemColors.ActiveCaption;
					button[i].Text = cat[i].Name;
					panel1.Controls.Add(this.button[i]);
					//機能をつける
					this.button[i].Click += new System.EventHandler(this.button_Click);
					Bqty++;
				}
			}
			catch { }

		}
		private void button_Click(object sender, EventArgs e)
		{
			try
			{
				//classを作る。
				if (panel2.Controls.Count > 0)
				{
					panel2.Controls.RemoveAt(0);
				}
				UserControl UC = dict[((System.Windows.Forms.ButtonBase)sender).Text];
				for (int i = 0; i < Bqty; i++)
				{
					button[i].BackColor = System.Drawing.SystemColors.ActiveCaption;
				}
				((System.Windows.Forms.ButtonBase)sender).BackColor = Color.FromArgb(121, 82, 178);

				panel2.Controls.Add(UC);
			}
			catch { MessageBox.Show("このボタンのアクセス権を持っていません"); }
		}

		private void button1_Click(object sender, EventArgs e)
		{
			int search = 1;
			button_setting(search);
			button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button3.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button5.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button1.BackColor = Color.FromArgb(121, 82, 178);
		}

		private void button2_Click(object sender, EventArgs e)
		{
			int search = 2;
			button_setting(search);
			button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button3.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button5.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button2.BackColor = Color.FromArgb(121, 82, 178);
		}

		private void button3_Click(object sender, EventArgs e)
		{
			int search = 3;
			button_setting(search);
			button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button5.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button3.BackColor = Color.FromArgb(121, 82, 178);
		}

		private void button5_Click(object sender, EventArgs e)
		{
			int search = 4;
			button_setting(search);
			button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button3.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			button5.BackColor = Color.FromArgb(121, 82, 178);
		}

		private void button7_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}
	}
}
