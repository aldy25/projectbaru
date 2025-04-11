using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using NPOI.SS.Formula.Functions;

namespace 給料計算アプリサンプル
{
	public partial class Re_03 : UserControl
	{
		private WebView2 webView;
		private Re_03 pdfViewer;

		public Re_03()
		{
			InitializeComponent();
			InitializePdfViewer();
		}
		private async void InitializePdfViewer()
		{
			webView = new WebView2
			{
				Dock = DockStyle.Fill
			};
			panel1.Controls.Add(webView);
			await webView.EnsureCoreWebView2Async();
			LoadPdf(@"C:\Users\dgghx\OneDrive\デスクトップ\OutputFile3.pdf");
		}

		// PDFを読み込むメソッド
		public void LoadPdf(string filePath)
		{
			if (System.IO.File.Exists(filePath))
			{
				webView.Source = new Uri($"file:///{filePath.Replace("\\", "/")}");
			}
			else
			{
				MessageBox.Show("PDFファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void Re_03_Load(object sender, EventArgs e)
		{
			
		}
		
	}
}
