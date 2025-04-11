using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using NPOI.SS.Formula.Functions;
using static 給料計算アプリサンプル.Userbase;

namespace 給料計算アプリサンプル
{
	internal class Re_03C : ReBase
	{
		List<Payroll> payList;
		DateTime StartDay;
		DateTime Endday;
		DateTime Payday;
		int WorkingDays;
		public Re_03C(List<Payroll> plist,DateTime startday, DateTime endday, DateTime payday, int workingdays)
		{
			payList = plist;
			StartDay = startday;
			Endday = endday;
			Payday = payday;
			WorkingDays = workingdays;
		}
		public void makeReport3()
		{
			//人数分業務をおこなう

			for (int i = 0; i < payList.Count; i++)
			{
				// PDFファイルの読み込み
				string foldername = FoldeNameGet();
				string templatePath = foldername + @"\\SalarySlip_Format.pdf"; // テンプレートファイルのパス
				PdfReader reader = new PdfReader(templatePath);

				// レイアウトの加工を行うためのPdfStamperを作成
				string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				string outputPath = desktopPath + @"\\OutputFile3.pdf";   // 出力先のパス
				PdfStamper stamper = new PdfStamper(reader, new FileStream(outputPath, FileMode.Create));

				// ページのコンテンツを取得
				PdfContentByte content = stamper.GetOverContent(1);

				// フォントの設定
				BaseFont font = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
				content.SetFontAndSize(font, 7);

				// テキストの描画
				content.BeginText();
				content.SetRGBColorFill(255, 0, 0);

				// MONTHS 給料日
				try
				{
					// 月名を英語で取得
					string monthName = Endday.ToString("MMMM", new System.Globalization.CultureInfo("en-US"));
					// 出力する文字列を作成
					string result = $"MONTHS : {monthName} {Endday:yyyy} ({StartDay:dd MMMM} - {Endday:dd MMMM yyyy})";
					content.ShowTextAligned(Element.ALIGN_LEFT, result, 115, 495, 0);
				}
				catch { }
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Name, 160, 485, 0);// Employee Name
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].P_code.ToString(), 160, 475, 0); // Employee Number
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Title, 160, 467, 0); // Position
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Dept, 220, 467, 0); // Department

				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Basic.ToString(), 184, 451, 0); // Basic Salary
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Allowance.ToString(), 184, 443, 0); // Allowance
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Transport.ToString(), 184, 436, 0); // Transportation
				long OT = (long)Math.Round(payList[i].Basic / WorkingDays / 8 * (payList[i].OTA * 1.5 + payList[i].OTB * 2 + payList[i].OTC * 3 + payList[i].OTD * 4), 0);
				content.ShowTextAligned(Element.ALIGN_LEFT, OT.ToString(), 184, 428, 0); // Overtime
				//content.ShowTextAligned(Element.ALIGN_LEFT, , 184, 421, 0); // Meal Allowances
				//content.ShowTextAligned(Element.ALIGN_LEFT, "@aaaa@", 184, 413, 0); // Overtime Meal Allowances

				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Other.ToString(), 184, 406, 0); // Others

				//content.ShowTextAligned(Element.ALIGN_LEFT, , 331, 451, 0); // Loan
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].BPJSkesh.ToString(), 331, 443, 0); // BPJS Kesehatan   
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].BPJSNakar.ToString(), 331, 436, 0); // BPJS Tenaga Kerja    
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].BPJSpen.ToString(), 331, 428, 0); // BPJS Pensiun
				//content.ShowTextAligned(Element.ALIGN_LEFT, "@gggg@", 331, 421, 0); // Ultratech
				//content.ShowTextAligned(Element.ALIGN_LEFT, "@hhhh@", 331, 413, 0); // Absence
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].SPcost.ToString(), 331, 406, 0); // Potongan Iuran SP
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Other.ToString(), 331, 398, 0); // Others

				long SubTotal = payList[i].Basic + OT + payList[i].Transport + payList[i].Allowance;
				content.ShowTextAligned(Element.ALIGN_LEFT, SubTotal.ToString(), 184, 365, 0); // Total Amount
				long SubTotal2 = payList[i].BPJSpen + payList[i].BPJSNakar + payList[i].BPJSkesh + payList[i].Other + payList[i].SPcost;
				content.ShowTextAligned(Element.ALIGN_LEFT, SubTotal2.ToString(), 331, 365, 0); // Total Deduction
				content.ShowTextAligned(Element.ALIGN_LEFT, (SubTotal-SubTotal2).ToString(), 331, 356, 0); // NET  INCOME
																					// Said amount
				string angka = Angka(SubTotal - SubTotal2);
				content.ShowTextAligned(Element.ALIGN_LEFT, angka, 120, 348, 0); 

				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Bank, 120, 340, 0); // Employee Account  Number
				content.ShowTextAligned(Element.ALIGN_LEFT, Payday.ToString("dd MMMM yyyy"), 120, 332, 0); // Date
				//content.ShowTextAligned(Element.ALIGN_LEFT, "@qqqq@", 120, 324, 0); // Note

				content.ShowTextAligned(Element.ALIGN_LEFT, "May Bank  – Deltamas", 294, 339, 0); // Name of Bank   
				content.ShowTextAligned(Element.ALIGN_LEFT, payList[i].Name, 294, 332, 0); // Employee  Signature  

				content.SetFontAndSize(font, 6);
				for (int i1 = 0; i1 < 31; i1++)
				{
					// 対象年月
					int year = 2025; // TODO: データから設定できるようにする
					int month = 1; // TODO: データから設定できるようにする
					DateTime date = new DateTime(year, month, 20).AddDays(i1);
					string ymd = date.ToString("dd-MMM-yy", CultureInfo.CreateSpecificCulture("en-US"));
					content.ShowTextAligned(Element.ALIGN_LEFT, ymd, 68, 283 - ((int)(i1 * 7.26)), 0); // Date
					string week = date.ToString("ddd", CultureInfo.CreateSpecificCulture("en-US"));
					content.ShowTextAligned(Element.ALIGN_LEFT, week, 108, 283 - ((int)(i1 * 7.26)), 0); // Day

					//Shift

					//OT Hour

					//Regulation
					//1.5倍
					//2.0倍
					//3.0倍
					//4.0倍
					//Total

				}
				content.EndText();

				// 変更を保存して終了
				stamper.Close();
				reader.Close();
			}
		}

		//送信

		//インドネシア語数字に切り替える
		public string Angka(long num)
		{
			string[] satuan = { "", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan" };
			string[] belasan = { "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas" };
			string[] puluhan = { "", "", "dua puluh", "tiga puluh", "empat puluh", "lima puluh", "enam puluh", "tujuh puluh", "delapan puluh", "sembilan puluh" };


			if (num == 0) return "nol";

			if (num < 10)
				return satuan[num];

			if (num < 20)
				return belasan[num - 10];

			if (num < 100)
				return puluhan[num / 10] + (num % 10 != 0 ? " " + Angka(num % 10) : "");

			if (num < 200)
				return "seratus" + (num % 100 != 0 ? " " + Angka(num % 100) : "");

			if (num < 1000)
				return satuan[num / 100] + " ratus" + (num % 100 != 0 ? " " + Angka(num % 100) : "");

			if (num < 2000)
				return "seribu" + (num % 1000 != 0 ? " " + Angka(num % 1000) : "");

			if (num < 1000000)
				return Angka(num / 1000) + " ribu" + (num % 1000 != 0 ? " " + Angka(num % 1000) : "");

			if (num < 1000000000)
				return Angka(num / 1000000) + " juta" + (num % 1000000 != 0 ? " " + Angka(num % 1000000) : "");

			if (num < 1000000000000)
				return Angka(num / 1000000000) + " miliar" + (num % 1000000000 != 0 ? " " + Angka(num % 1000000000) : "");

			return "Angka terlalu besar";
		}
	}
}
