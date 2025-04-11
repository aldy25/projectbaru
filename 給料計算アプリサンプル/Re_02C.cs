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
	internal class Re_02C : ReBase
	{
		List<Payroll> payList;
		int workingDays;
		public Re_02C(List<Payroll> plist, int WorkingDays)
		{
			payList = plist;
			workingDays = WorkingDays;
		}

		//②JACデータの作成
		public void makeReport2()
		{
			string foldername = FoldeNameGet();
			string templatePath = foldername + @"\\AccountingSalaryDocument_Format2.xlsx"; // テンプレートファイルのパス

			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			string outputPath = desktopPath + @"\\OutputFile2.xlsx";   // 出力先のパス

			if (!File.Exists(templatePath))
			{
				Console.WriteLine("テンプレートファイルが見つかりません。");
				return;
			}
			//Meal Allowanceデータ取得
			long MealAllowance = 0;
			try
			{
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
			}
			catch { }
			

			try
			{
				
				// テンプレートを開く
				using (FileStream fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
				{
					long GrandTotal = 0;
					IWorkbook workbook = new XSSFWorkbook(fs);
					ISheet sheet = workbook.GetSheetAt(0); // 最初のシートを取得

					// データを入力
					for (int i = 0; i < payList.Count; i++)
					{
						//No
						IRow row = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell = row.GetCell(0) ?? row.CreateCell(0);
						cell.SetCellValue(i + 1);

						//Employee ID
						IRow row1 = sheet.GetRow(1+ i) ?? sheet.CreateRow(1 + i);
						ICell cell1 = row1.GetCell(1) ?? row1.CreateCell(1);
						cell1.SetCellValue(payList[i].P_code);

						//Name 
						IRow row2 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell2 = row2.GetCell(2) ?? row2.CreateCell(2);
						cell2.SetCellValue(payList[i].Name);

						//Department
						IRow row3 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell3 = row3.GetCell(3) ?? row3.CreateCell(3);
						cell3.SetCellValue(payList[i].Dept);

						//Basic Salary
						IRow row4 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell4 = row4.GetCell(4) ?? row4.CreateCell(4);
						cell4.SetCellValue(payList[i].Basic);
						GrandTotal = GrandTotal + payList[i].Basic;

						//Allowance
						IRow row5 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell5 = row5.GetCell(5) ?? row5.CreateCell(5);
						cell5.SetCellValue(payList[i].Allowance);
						GrandTotal = GrandTotal + payList[i].Allowance;

						//Over Time
						IRow row6 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell6 = row6.GetCell(6) ?? row6.CreateCell(6);
						//***注意***時給の出し方を確認する。
						long OT = (long)Math.Round(payList[i].Basic / workingDays / 8 * (payList[i].OTA * 1.5 + payList[i].OTB * 2 + payList[i].OTC * 3 + payList[i].OTD * 4), 0);
						cell6.SetCellValue(OT);
						GrandTotal = GrandTotal + OT;

						//Transport***未完***
						IRow row7 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell7 = row7.GetCell(7) ?? row7.CreateCell(7);
						cell7.SetCellValue(payList[i].Transport);
						GrandTotal = GrandTotal + payList[i].Transport;

						//Meal Allowance ***未完***
						IRow row8 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell8 = row8.GetCell(8) ?? row8.CreateCell(8);
						cell8.SetCellValue(MealAllowance);
						GrandTotal = GrandTotal + MealAllowance;

						//Overtime Meal Allowances 残業食事手当　***未完***
						//IRow row9 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						//ICell cell9 = row9.GetCell(9) ?? row9.CreateCell(9);
						//cell9.SetCellValue(MealAllowance);
						//GrandTotal = GrandTotal + ;

						//Rapelan 未払い給料 ***未完***
						//IRow row10 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						//ICell cell10 = row10.GetCell(10) ?? row10.CreateCell(10);
						//cell10.SetCellValue(MealAllowance);
						//GrandTotal = GrandTotal + ;

						//Adj Penambahan 追加 ***未完***
						//IRow row11 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						//ICell cell11 = row11.GetCell(11) ?? row11.CreateCell(11);
						//cell11.SetCellValue(MealAllowance);
						//GrandTotal = GrandTotal + ;

						//IURAN SP 組合費
						IRow row12 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell12 = row12.GetCell(12) ?? row12.CreateCell(12);
						cell12.SetCellValue(payList[i].SPcost);
						GrandTotal = GrandTotal - payList[i].SPcost;

						//BPJS Pensiun
						IRow row13 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell13 = row13.GetCell(13) ?? row13.CreateCell(13);
						cell13.SetCellValue(payList[i].BPJSpen);
						GrandTotal = GrandTotal - payList[i].BPJSpen;

						//BPJS TK
						IRow row14 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell14 = row14.GetCell(14) ?? row14.CreateCell(14);
						cell14.SetCellValue(payList[i].BPJSNakar);
						GrandTotal = GrandTotal - payList[i].BPJSNakar;

						//BPJS Kesehatan
						IRow row15 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell15 = row15.GetCell(15) ?? row15.CreateCell(15);
						cell15.SetCellValue(payList[i].BPJSkesh);
						GrandTotal = GrandTotal - payList[i].BPJSkesh;

						//Other Loan
						IRow row16 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell16 = row16.GetCell(16) ?? row16.CreateCell(16);
						cell16.SetCellValue(payList[i].Other);
						GrandTotal = GrandTotal - payList[i].Other;

						//給料前借***未完***
						//IRow row17 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						//ICell cell17 = row17.GetCell(17) ?? row17.CreateCell(17);
						//cell17.SetCellValue(payList[i].Other);
						//GrandTotal = GrandTotal -;

						//Absence 欠席による給料差引
						//IRow row18 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						//ICell cell18 = row18.GetCell(18) ?? row18.CreateCell(18);
						//cell18.SetCellValue(payList[i].Other);
						//GrandTotal = GrandTotal -;

						//Adj(Pengurangan)追加減算
						//IRow row19 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						//ICell cell19 = row19.GetCell(19) ?? row16.CreateCell(19);
						//cell19.SetCellValue(payList[i].Other);
						//GrandTotal = GrandTotal -;

						//pay 支給額
						IRow row20 = sheet.GetRow(1 + i) ?? sheet.CreateRow(1 + i);
						ICell cell20 = row20.GetCell(20) ?? row20.CreateCell(20);
						cell20.SetCellValue(GrandTotal);

					}
					// 新しいExcelファイルとして保存
					using (FileStream outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
					{
						workbook.Write(outFs);
					}
				}
				MessageBox.Show("Excelファイルを作成しました: " + outputPath);
			}
			catch (Exception ex)
			{
				Console.WriteLine("エラー: " + ex.Message);
			}
		}


	}
}
