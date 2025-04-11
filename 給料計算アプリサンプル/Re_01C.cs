using Npgsql;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static NPOI.HSSF.Util.HSSFColor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using static 給料計算アプリサンプル.Userbase;

namespace 給料計算アプリサンプル
{
	internal class Re_01C : ReBase
	{
		List<Payroll> payList;
		int workingDays;
		public Re_01C(List<Payroll> plist,int WorkingDays)
		{
			payList=plist;
			workingDays = WorkingDays;
		}

		//①給料振込データの作成
		
		public void makeReport1()
		{
			string foldername = FoldeNameGet();
			string templatePath = foldername+ @"\\AccountingSalaryDocument_Format.xlsx"; // テンプレートファイルのパス

			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			string outputPath = desktopPath+@"\\OutputFile.xlsx";   // 出力先のパス

			if (!File.Exists(templatePath))
			{
				Console.WriteLine("テンプレートファイルが見つかりません。");
				return;
			}
			//Meal Allowanceデータ取得
			long MealAllowance = 0;
			using (var conn = new NpgsqlConnection(connectionString))
			{
				conn.Open();
				string sql = @"
				SELECT c_mealallowance 
				FROM company 
				WHERE c_update <= CURRENT_DATE 
				ORDER BY c_update DESC 
				LIMIT 1";

				using (var cmd = new NpgsqlCommand(sql, conn))
				{
					using (var reader = cmd.ExecuteReader())
					{
						if (reader.Read()) // データがある場合
						{
							MealAllowance = reader.GetInt64(0);
						}
					}
				}
			}

			try
			{
				// テンプレートを開く
				var workbook = WorkbookFactory.Create(templatePath);
				//using (FileStream fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
				//{
					//IWorkbook workbook = new XSSFWorkbook(fs);
					ISheet sheet = workbook.GetSheetAt(0); // 最初のシートを取得


					// データを入力
					for (int i = 0; i < payList.Count; i++)
					{
						//No
						IRow row = sheet.GetRow(3+i) ?? sheet.CreateRow(3 + i);
						ICell cell = row.GetCell(1) ?? row.CreateCell(1);
						cell.SetCellValue(i+1);

						//Employee ID
						IRow row1 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell1 = row1.GetCell(2) ?? row1.CreateCell(2);
						cell1.SetCellValue(payList[i].P_code);

						//Name 
						IRow row2 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell2 = row2.GetCell(3) ?? row2.CreateCell(3);
						cell2.SetCellValue(payList[i].Name);

						//Bank Account 
						IRow row3 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell3 = row3.GetCell(4) ?? row3.CreateCell(4);
						cell3.SetCellValue(payList[i].Bank);

						//Basic Salary
						IRow row4 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell4 = row4.GetCell(5) ?? row4.CreateCell(5);
						cell4.SetCellValue(payList[i].Basic );

						//Transport***未完***
						IRow row5 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell5 = row5.GetCell(6) ?? row5.CreateCell(6);
						cell5.SetCellValue(payList[i].Transport);

						//Over Time
						IRow row6 = sheet.GetRow(3+ i) ?? sheet.CreateRow(3 + i);
						ICell cell6 = row6.GetCell(8) ?? row6.CreateCell(8);
						//***注意***時給の出し方を確認する。
						long OT = (long)Math.Round(payList[i].Basic/ workingDays/ 8 * (payList[i].OTA*1.5+ payList[i].OTB*2+ payList[i].OTC * 3 + payList[i].OTD*4), 0);
						cell6.SetCellValue(OT);

						//Allowance
						IRow row7 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell7 = row7.GetCell(9) ?? row7.CreateCell(9);
						cell7.SetCellValue(payList[i].Allowance);

						//Meal Allowance ***未完***
						IRow row8 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell8 = row8.GetCell(11) ?? row8.CreateCell(11);
						cell8.SetCellValue(MealAllowance);

						//sub Total
						IRow row9 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell9 = row9.GetCell(13) ?? row9.CreateCell(13);
						long SubTotal = payList[i].Basic + OT + payList[i].Transport + payList[i].Allowance + MealAllowance;
						cell9.SetCellValue(SubTotal);

						//BPJS Pensiun
						IRow row10 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell10 = row10.GetCell(14) ?? row10.CreateCell(14);
						cell10.SetCellValue(payList[i].BPJSpen);

						//BPJS TK
						IRow row11 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell11 = row11.GetCell(15) ?? row11.CreateCell(15);
						cell11.SetCellValue(payList[i].BPJSNakar);

						//BPJS Kesehatan
						IRow row12 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell12 = row12.GetCell(16) ?? row12.CreateCell(16);
						cell12.SetCellValue(payList[i].BPJSkesh);

						//Other Loan
						IRow row13 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell13 = row13.GetCell(18) ?? row13.CreateCell(18);
						cell13.SetCellValue(payList[i].Other);

						//IURAN SP 組合費
						IRow row14 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell14 = row14.GetCell(20) ?? row14.CreateCell(20);
						cell14.SetCellValue(payList[i].SPcost);

						//sub Total2 天引き総額
						long SubTotal2 = 0;
						SubTotal2 = payList[i].BPJSpen + payList[i].BPJSNakar + payList[i].BPJSkesh + payList[i].Other + payList[i].SPcost;
						IRow row15 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell15 = row15.GetCell(22) ?? row15.CreateCell(22);
						cell15.SetCellValue(SubTotal2);

						//pay 支給額
						IRow row16 = sheet.GetRow(3 + i) ?? sheet.CreateRow(3 + i);
						ICell cell16 = row16.GetCell(23) ?? row16.CreateCell(23);
						cell16.SetCellValue(SubTotal-SubTotal2);

					}
				// 新しいExcelファイルとして保存
				using (MemoryStream ms = new MemoryStream())
				{
					workbook.Write(ms);
					using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
					{
						byte[] data1 = ms.ToArray();
						fs.Write(data1, 0, data1.Length);
						fs.Flush();
					}
					workbook = null;
				}	
				//}
				MessageBox.Show("Excelファイルを作成しました: " + outputPath);
			}
			catch (Exception ex)
			{
				Console.WriteLine("エラー: " + ex.Message);
			}

			

			
		}
	}
}
