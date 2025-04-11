using Npgsql;
using Org.BouncyCastle.Asn1.IsisMtt.X509;
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

namespace 給料計算アプリサンプル
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//☆パスワードが一致するか確認しログインを許可するかどうか決める。
			//①部門名が選択されているかどうかを確認する。
			//②パスワードのデータをpasswordから引き出す。
			string username = ConfigurationManager.AppSettings["Username"];
			string password = ConfigurationManager.AppSettings["Password"];

			try
			{
				//if (username == textBox1.Text && password == textBox2.Text)
				//{
					P_01 fm2 = new P_01();
					fm2.Show();
					this.Hide();
				//}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Kata sandi yang berbeda. " + ex.Message);
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			
		}
	}
}
