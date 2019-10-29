using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace PgmDataMonitor
{
    public partial class FormConfirmWindow : Form
    {
		string title = "";
		
        public FormConfirmWindow()
        {
            InitializeComponent();
			listView1.Width = Form1.SETDATA.listviewWidth;
			listView1.Height = Form1.SETDATA.listviewHeight;
        }

        private void FormConfirmWindow_Load(object sender, EventArgs e)
        {
			this.Text = title;
        }

		public void SetTarget(string tar_file)
		{
			//成型機を判定
			bool isMulti = false;
			bool isLS_NQD = false;
			bool isHS = false;
			if(tar_file.IndexOf("_マルチCav") >= 0)//マルチか？
			{
				isMulti = true;
			}
			if(tar_file.IndexOf("_LS") >= 0)
			{
				isLS_NQD = true;
			}
			else if(tar_file.IndexOf("_NQD") >= 0)
			{
				isLS_NQD = true;
			}
			else if(tar_file.IndexOf("_HS") >= 0)
			{
				isHS = true;
			}

			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			if(tar_file.IndexOf("_monthly") >= 0)
			{
				stCurrentDir = stCurrentDir + "\\monthly";
				title = tar_file.Substring(0, 7);
			}
			else
			{
				title = tar_file.Substring(7, 7);
			}

            string currentCsvFile = stCurrentDir + "\\" + tar_file;

			string machineName = "";
			string ratioOfHousha = "";
			string opratorName = "";

			this.ActiveControl = this.listView1;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.View = View.Details;
            listView1.HideSelection = false;

            //全体の色設定
            listView1.Sorting = SortOrder.None;
            listView1.ForeColor = Color.Black;//初期の色
            listView1.BackColor = Color.Coral;//全体背景色

			if(isLS_NQD)//LS or NQD
			{
                string shimeSign = "";
				string sleeveNo = "";
				string Tkeisu = "";
				string z3hosei = "";
				string z3Value = "";
				string ct1Value = "";
				string ct2Value = "";
				string cc32Value = "";
				string cpValue = "";
				string tactValue = "";
				string dateValue = "";
				string timeValue = "";
				string nikuUpLimit = "";
				string nikuLoLimit = "";
				string shotCnt = "";
				string shukaiCnt = "";

				if(!isMulti)//ｼﾝｸﾞﾙLS or NQD
				{
					string nikuData = "";
					string nikuResult = "";
					string resultCause = "";

		            ColumnHeader columnShime;
		            ColumnHeader columnSleeve;
		            ColumnHeader columnTkeisu;
		            ColumnHeader columnInitialTemp;
		            ColumnHeader columnSeikeiTemp;
		            ColumnHeader columnKaatsuTime;
		            ColumnHeader columnZ3Value;
		            ColumnHeader columnZ3hosei;
		            ColumnHeader columnCpValue;
		            ColumnHeader columnNikuatsuValue;
		            ColumnHeader columnNikuatsuUpper;
		            ColumnHeader columnNikuatsuLower;
		            ColumnHeader columnTact;
		            ColumnHeader columnDate;
		            ColumnHeader columnTimeStamp;
		            ColumnHeader columnResult;
		            ColumnHeader columnNgCause;
		            ColumnHeader columnShotCount;
		            ColumnHeader columnSeikeiCount;

					columnShime = new ColumnHeader();
		            columnSleeve = new ColumnHeader();
		            columnTkeisu = new ColumnHeader();
		            columnInitialTemp = new ColumnHeader();
		            columnSeikeiTemp = new ColumnHeader();
		            columnKaatsuTime = new ColumnHeader();
		            columnZ3Value = new ColumnHeader();
		            columnZ3hosei = new ColumnHeader();
		            columnCpValue = new ColumnHeader();
		            columnNikuatsuValue = new ColumnHeader();
		            columnNikuatsuUpper = new ColumnHeader();
		            columnNikuatsuLower = new ColumnHeader();
		            columnTact = new ColumnHeader();
		            columnTimeStamp = new ColumnHeader();
					columnDate = new ColumnHeader();
					columnResult = new ColumnHeader();
					columnNgCause = new ColumnHeader();
					columnShotCount = new ColumnHeader();
					columnSeikeiCount = new ColumnHeader();

					columnShime.Text = "締め";
		            columnSleeve.Text = "SL";
		            columnTkeisu.Text = "T係数";
		            columnInitialTemp.Text = "初期温度";
		            columnSeikeiTemp.Text = "成型温度";
		            columnKaatsuTime.Text = "加圧時間";
		            columnZ3Value.Text = "Z3";
		            columnZ3hosei.Text = "Z3補正";
		            columnCpValue.Text = "CP値";
		            columnTact.Text = "ﾀｸﾄ";
					columnDate.Text = "日付";
		            columnTimeStamp.Text = "時刻";
		            columnNikuatsuUpper.Text = "肉厚上限";
		            columnNikuatsuValue.Text = "測定値";
		            columnNikuatsuLower.Text = "肉厚下限";
		            columnResult.Text = "合否";
		            columnNgCause.Text = "原因";
		            columnShotCount.Text = "ｼｮｯﾄ数";
		            columnSeikeiCount.Text = "成型数";

		            ColumnHeader[] colHeaderRegValue =
		              { columnShime, columnSleeve, columnTkeisu, columnZ3hosei, columnZ3Value, columnInitialTemp, columnSeikeiTemp, columnKaatsuTime, 
		              columnCpValue, columnTact, columnDate, columnTimeStamp, columnNikuatsuUpper, columnNikuatsuValue, columnNikuatsuLower, columnResult, columnNgCause, columnShotCount, columnSeikeiCount};
		            listView1.Columns.AddRange(colHeaderRegValue);

					//CSVファイルを読み込む
		            var readToEnd = File.ReadAllLines(currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
		            int lines = readToEnd.Length;

		            for (int i = 0; i < lines; i++)
		            {
						if(i == 0)//ヘッダ部はスキップ
						{
							continue;
						}
		                //１行のstringをstream化してTextFieldParserで処理する
		                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
		                {
		                    using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
		                    {
		                        parser.TextFieldType = FieldType.Delimited;
		                        parser.Delimiters = new[] { "," };
		                        parser.HasFieldsEnclosedInQuotes = true;
		                        parser.TrimWhiteSpace = false;
		                        string[] fields = parser.ReadFields();

		                        for (int j = 0; j < fields.Length; j++)
		                        {
									if(j == 0)//date
									{
										dateValue = fields[j];
									}
									else if(j == 4)//ログファイル名
									{
									}
									else if(j == 20)//time
									{
										timeValue = fields[j];
										if(timeValue.Length == 19)
										{
											timeValue = fields[j].Substring(11, 8);
											dateValue = fields[j].Substring(0, 10);
										}
									}
									else if(j == 2)//seikeikiName
									{
										machineName = fields[j];
									}
									else if(j == 3)//seihinName
									{
									}
									else if(j == 4)//lastFileName
									{
									}
									else if(j == 5)//shimeSign
									{
										shimeSign = fields[j];
									}
									else if(j == 6)//sleeveNo
									{
										sleeveNo = fields[j];
									}
									else if(j == 7)//z3Value
									{
										z3Value = fields[j];
									}
									else if(j == 8)//ct1Value
									{
										ct1Value = fields[j];
										string r = ct1Value.Replace("℃", "C");
										ct1Value = r;
									}
									else if(j == 9)//ct2Value
									{
										ct2Value = fields[j];
										string r = ct2Value.Replace("℃", "C");
										ct2Value = r;
									}
									else if(j == 13)//cc32Value
									{
										cc32Value = fields[j];
									}
									else if(j == 14)//cpValue
									{
										cpValue = fields[j];
									}
									else if(j == 15)//tactValue
									{
										tactValue = fields[j];
									}
									else if(j == 16)//Tkeisu
									{
										Tkeisu = fields[j];
									}
									else if(j == 17)//z3hosei
									{
										z3hosei = fields[j];
									}
									else if(j == 21)//nikuUpLimit
									{
										nikuUpLimit = fields[j];
									}
									else if(j == 22)//nikuData
									{
										nikuData = fields[j];
									}
									else if(j == 23)//nikuLoLimit
									{
										nikuLoLimit = fields[j];
									}
									else if(j == 24)//nikuResult
									{
										nikuResult = fields[j];
									}
									else if(j == 25)//resultCause
									{
										resultCause = fields[j];
									}
									else if(j == 26)//currentHousharitsu
									{
										ratioOfHousha = fields[j];
									}
									else if(j == 27)//currentOperator
									{
										opratorName = fields[j];
									}
									else if(j == 28)//shotCount
									{
										shotCnt = fields[j];
									}
									else if(j == 29)//seikeiCount
									{
										shukaiCnt = fields[j];
									}
		                        }
		                    }
		                    
		                }

						string[] item1 = {shimeSign, sleeveNo, Tkeisu, z3hosei, z3Value, ct1Value, ct2Value, cc32Value, 
						cpValue, tactValue, dateValue, timeValue, 
						nikuUpLimit, nikuData, nikuLoLimit, nikuResult, resultCause, shotCnt, shukaiCnt};

						listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
			            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);

						if(nikuResult == "OK")
						{
				            listView1.Items[0].BackColor = Color.Lime;
						}
						else if(nikuResult == "NG")
						{
				            listView1.Items[0].BackColor = Color.Red;
						}

						if(sleeveNo == "0" || sleeveNo == "" || sleeveNo == "D" || sleeveNo == "SD" || sleeveNo == "先" || sleeveNo == "空")
						{
							listView1.Items[0].BackColor = Color.Gray;
						}

						//日勤時間帯
						string strNoonSta = "08:00:00";
						string strNoonEnd = "17:00:00";
						DateTime noonSta = DateTime.Parse(strNoonSta);
						DateTime noonEnd = DateTime.Parse(strNoonEnd);

						//夕勤時間帯
						string strSunsetSta = "17:00:00";
						string strSunsetEnd = "23:59:59";
						DateTime sunsetSta = DateTime.Parse(strSunsetSta);
						DateTime sunsetEnd = DateTime.Parse(strSunsetEnd);

						//夜勤時間帯
						string strNightSta = "00:00:01";
						string strNightEnd = "08:00:00";
						DateTime nightSta = DateTime.Parse(strNightSta);
						DateTime nightEnd = DateTime.Parse(strNightEnd);

						DateTime dt3 = DateTime.Parse(timeValue);//タイムスタンプ(文字列)→DateTimeに変換

						if(noonSta <= dt3 && dt3 < noonEnd)
						{
							listView1.Items[0].ForeColor = Color.Blue;//青
						}
						if(sunsetSta <= dt3 && dt3 < sunsetEnd)
						{
							listView1.Items[0].ForeColor = Color.Green;//緑
						}
						if(nightSta <= dt3 && dt3 < nightEnd)
						{
							listView1.Items[0].ForeColor = Color.Purple;//赤
						}
		            }

		            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
				}
				else//ﾏﾙﾁLS or NQD
				{
					string [] nikuData = new string[] {"", "", "", "", "", ""};
					string [] nikuResult = new string[] {"", "", "", "", "", ""};
					string [] resultCause = new string[] {"", "", "", "", "", ""};

		            ColumnHeader columnShime;
		            ColumnHeader columnSleeve;
		            ColumnHeader columnTkeisu;
		            ColumnHeader columnInitialTemp;
		            ColumnHeader columnSeikeiTemp;
		            ColumnHeader columnKaatsuTime;
		            ColumnHeader columnZ3Value;
		            ColumnHeader columnZ3hosei;
		            ColumnHeader columnCpValue;
		            ColumnHeader columnNikuatsuValue1;
		            ColumnHeader columnNikuatsuValue2;
		            ColumnHeader columnNikuatsuValue3;
		            ColumnHeader columnNikuatsuValue4;
		            ColumnHeader columnNikuatsuValue5;
		            ColumnHeader columnNikuatsuValue6;
		            ColumnHeader columnResult1;
		            ColumnHeader columnResult2;
		            ColumnHeader columnResult3;
		            ColumnHeader columnResult4;
		            ColumnHeader columnResult5;
		            ColumnHeader columnResult6;
		            ColumnHeader columnNgCause1;
		            ColumnHeader columnNgCause2;
		            ColumnHeader columnNgCause3;
		            ColumnHeader columnNgCause4;
		            ColumnHeader columnNgCause5;
		            ColumnHeader columnNgCause6;
		            ColumnHeader columnNikuatsuUpper;
		            ColumnHeader columnNikuatsuLower;
		            ColumnHeader columnTact;
		            ColumnHeader columnDate;
		            ColumnHeader columnTimeStamp;
		            ColumnHeader columnShotCount;
		            ColumnHeader columnSeikeiCount;

		            columnShime = new ColumnHeader();
		            columnSleeve = new ColumnHeader();
		            columnTkeisu = new ColumnHeader();
		            columnInitialTemp = new ColumnHeader();
		            columnSeikeiTemp = new ColumnHeader();
		            columnKaatsuTime = new ColumnHeader();
		            columnZ3Value = new ColumnHeader();
		            columnZ3hosei = new ColumnHeader();
		            columnCpValue = new ColumnHeader();
		            columnNikuatsuValue1 = new ColumnHeader();
		            columnNikuatsuValue2 = new ColumnHeader();
		            columnNikuatsuValue3 = new ColumnHeader();
		            columnNikuatsuValue4 = new ColumnHeader();
		            columnNikuatsuValue5 = new ColumnHeader();
		            columnNikuatsuValue6 = new ColumnHeader();
		            columnNikuatsuUpper = new ColumnHeader();
		            columnNikuatsuLower = new ColumnHeader();
		            columnTact = new ColumnHeader();
		            columnTimeStamp = new ColumnHeader();
					columnDate = new ColumnHeader();
					columnResult1 = new ColumnHeader();
					columnResult2 = new ColumnHeader();
					columnResult3 = new ColumnHeader();
					columnResult4 = new ColumnHeader();
					columnResult5 = new ColumnHeader();
					columnResult6 = new ColumnHeader();
					columnNgCause1 = new ColumnHeader();
					columnNgCause2 = new ColumnHeader();
					columnNgCause3 = new ColumnHeader();
					columnNgCause4 = new ColumnHeader();
					columnNgCause5 = new ColumnHeader();
					columnNgCause6 = new ColumnHeader();
					columnShotCount = new ColumnHeader();
					columnSeikeiCount = new ColumnHeader();

					columnShime.Text = "締め";
		            columnSleeve.Text = "SL";
		            columnTkeisu.Text = "T係数";
		            columnInitialTemp.Text = "初期温度";
		            columnSeikeiTemp.Text = "成型温度";
		            columnKaatsuTime.Text = "加圧時間";
		            columnZ3Value.Text = "Z3";
		            columnZ3hosei.Text = "Z3補正";
		            columnCpValue.Text = "CP値";
		            columnTact.Text = "ﾀｸﾄ";
					columnDate.Text = "日付";
		            columnTimeStamp.Text = "時刻";
		            columnNikuatsuUpper.Text = "上限";
		            columnNikuatsuLower.Text = "下限";
		            columnNikuatsuValue1.Text = "Cav1";
		            columnResult1.Text = "Cav1合否";
		            columnNgCause1.Text = "Cav1原因";
		            columnNikuatsuValue2.Text = "Cav2";
		            columnResult2.Text = "Cav2合否";
		            columnNgCause2.Text = "Cav2原因";
		            columnNikuatsuValue3.Text = "Cav3";
		            columnResult3.Text = "Cav3合否";
		            columnNgCause3.Text = "Cav3原因";
		            columnNikuatsuValue4.Text = "Cav4";
		            columnResult4.Text = "Cav4合否";
		            columnNgCause4.Text = "Cav4原因";
		            columnNikuatsuValue5.Text = "Cav5";
		            columnResult5.Text = "Cav5合否";
		            columnNgCause5.Text = "Cav5原因";
		            columnNikuatsuValue6.Text = "Cav6";
		            columnResult6.Text = "Cav6合否";
		            columnNgCause6.Text = "Cav6原因";
		            columnShotCount.Text = "ｼｮｯﾄ数";
		            columnSeikeiCount.Text = "成型数";

		            ColumnHeader[] colHeaderRegValue =
		              { columnShime, columnSleeve, columnTkeisu, columnZ3hosei, columnZ3Value, columnInitialTemp, columnSeikeiTemp, columnKaatsuTime, 
		              columnCpValue, columnTact, columnDate, columnTimeStamp, columnNikuatsuUpper, columnNikuatsuLower, 
		              columnNikuatsuValue1, columnResult1, columnNgCause1, 
		              columnNikuatsuValue2, columnResult2, columnNgCause2, 
		              columnNikuatsuValue3, columnResult3, columnNgCause3, 
		              columnNikuatsuValue4, columnResult4, columnNgCause4, 
		              columnNikuatsuValue5, columnResult5, columnNgCause5, 
		              columnNikuatsuValue6, columnResult6, columnNgCause6, 
		              columnShotCount, columnSeikeiCount};
		            listView1.Columns.AddRange(colHeaderRegValue);

		            //CSV読込。高速化を狙い、最大行数を取得後、for分でループする
		            var readToEnd = File.ReadAllLines(currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
		            int lines = readToEnd.Length;

		            for (int i = 0; i < lines; i++)
		            {
						if(i == 0)//ヘッダ部はスキップ
						{
							continue;
						}
		                //１行のstringをstream化してTextFieldParserで処理する
		                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
		                {
		                    using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
		                    {
		                        parser.TextFieldType = FieldType.Delimited;
		                        parser.Delimiters = new[] { "," };
		                        parser.HasFieldsEnclosedInQuotes = true;
		                        parser.TrimWhiteSpace = false;
		                        string[] fields = parser.ReadFields();

		                        for (int j = 0; j < fields.Length; j++)
		                        {
									if(j == 0)//date
									{
										dateValue = fields[j];
									}
									else if(j == 4)//ログファイル名
									{
									}
									else if(j == 20)//time
									{
										timeValue = fields[j];
										if(timeValue.Length == 19)
										{
											timeValue = fields[j].Substring(11, 8);
											dateValue = fields[j].Substring(0, 10);
										}
									}
									else if(j == 2)//seikeikiName
									{
										machineName = fields[j];
									}
									else if(j == 3)//seihinName
									{
									}
									else if(j == 4)//lastFileName
									{
									}
									else if(j == 5)//shimeSign
									{
										shimeSign = fields[j];
									}
									else if(j == 6)//sleeveNo
									{
	                                    sleeveNo = fields[j];
									}
									else if(j == 7)//z3Value
									{
	                                    z3Value = fields[j];
									}
									else if(j == 8)//ct1Value
									{
	                                    ct1Value = fields[j];
										string r = ct1Value.Replace("℃", "C");
										ct1Value = r;
									}
									else if(j == 9)//ct2Value
									{
	                                    ct2Value = fields[j];
										string r = ct2Value.Replace("℃", "C");
										ct2Value = r;
									}
									else if(j == 13)//cc32Value
									{
	                                    cc32Value = fields[j];
									}
									else if(j == 14)//cpValue
									{
	                                    cpValue = fields[j];
									}
									else if(j == 15)//tactValue
									{
	                                    tactValue = fields[j];
									}
									else if(j == 16)//Tkeisu
									{
	                                    Tkeisu = fields[j];
									}
									else if(j == 17)//z3hosei
									{
	                                    z3hosei = fields[j];
									}
									else if(j == 21)//nikuUpLimit
									{
										nikuUpLimit = fields[j];
									}
									else if(j == 22)//nikuLoLimit
									{
										nikuLoLimit = fields[j];
									}

									else if(j == 23)//nikuData(Cav1)
									{
										nikuData[0] = fields[j];
									}
									else if(j == 26)//nikuData(Cav2)
									{
										nikuData[1] = fields[j];
									}
									else if(j == 29)//nikuData(Cav3)
									{
										nikuData[2] = fields[j];
									}
									else if(j == 32)//nikuData(Cav4)
									{
										nikuData[3] = fields[j];
									}
									else if(j == 35)//nikuData(Cav5)
									{
										nikuData[4] = fields[j];
									}
									else if(j == 38)//nikuData(Cav6)
									{
										nikuData[5] = fields[j];
									}

									else if(j == 24)//nikuResult(Cav1)
									{
										nikuResult[0] = fields[j];
									}
									else if(j == 27)//nikuResult(Cav2)
									{
										nikuResult[1] = fields[j];
									}
									else if(j == 30)//nikuResult(Cav3)
									{
										nikuResult[2] = fields[j];
									}
									else if(j == 33)//nikuResult(Cav4)
									{
										nikuResult[3] = fields[j];
									}
									else if(j == 36)//nikuResult(Cav5)
									{
										nikuResult[4] = fields[j];
									}
									else if(j == 39)//nikuResult(Cav6)
									{
										nikuResult[5] = fields[j];
									}

									else if(j == 25)//resultCause(Cav1)
									{
										resultCause[0] = fields[j];
									}
									else if(j == 28)//resultCause(Cav2)
									{
										resultCause[1] = fields[j];
									}
									else if(j == 31)//resultCause(Cav3)
									{
										resultCause[2] = fields[j];
									}
									else if(j == 34)//resultCause(Cav4)
									{
										resultCause[3] = fields[j];
									}
									else if(j == 37)//resultCause(Cav5)
									{
										resultCause[4] = fields[j];
									}
									else if(j == 40)//resultCause(Cav6)
									{
										resultCause[5] = fields[j];
									}

									else if(j == 41)//currentHousharitsu
									{
										ratioOfHousha = fields[j];
									}
									else if(j == 42)//currentOperator
									{
										opratorName = fields[j];
									}
									else if(j == 43)//shotCount
									{
	                                    shotCnt = fields[j];
									}
									else if(j == 44)//seikeiCount
									{
	                                    shukaiCnt = fields[j];
									}
		                        }
		                    }
		                    
		                }

						string[] item1 = {shimeSign, sleeveNo, Tkeisu, z3hosei, z3Value, ct1Value, ct2Value, cc32Value, 
						cpValue, tactValue, dateValue, timeValue, 
						nikuUpLimit, nikuLoLimit, 
						nikuData[0], nikuResult[0], resultCause[0], 
						nikuData[1], nikuResult[1], resultCause[1], 
						nikuData[2], nikuResult[2], resultCause[2], 
						nikuData[3], nikuResult[3], resultCause[3], 
						nikuData[4], nikuResult[4], resultCause[4], 
						nikuData[5], nikuResult[5], resultCause[5], 
						shotCnt, shukaiCnt};

						listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
			            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);

						if(sleeveNo == "0" || sleeveNo == "" || sleeveNo == "D" || sleeveNo == "SD" || sleeveNo == "先" || sleeveNo == "空")
						{
							listView1.Items[0].BackColor = Color.Gray;
						}

						//日勤時間帯
						string strNoonSta = "08:00:00";
						string strNoonEnd = "17:00:00";
						DateTime noonSta = DateTime.Parse(strNoonSta);
						DateTime noonEnd = DateTime.Parse(strNoonEnd);

						//夕勤時間帯
						string strSunsetSta = "17:00:00";
						string strSunsetEnd = "23:59:59";
						DateTime sunsetSta = DateTime.Parse(strSunsetSta);
						DateTime sunsetEnd = DateTime.Parse(strSunsetEnd);

						//夜勤時間帯
						string strNightSta = "00:00:01";
						string strNightEnd = "08:00:00";
						DateTime nightSta = DateTime.Parse(strNightSta);
						DateTime nightEnd = DateTime.Parse(strNightEnd);

						DateTime dt3 = DateTime.Parse(timeValue);//タイムスタンプ(文字列)→DateTimeに変換

						if(noonSta <= dt3 && dt3 < noonEnd)
						{
							listView1.Items[0].ForeColor = Color.Blue;//青
						}
						if(sunsetSta <= dt3 && dt3 < sunsetEnd)
						{
							listView1.Items[0].ForeColor = Color.Green;//緑
						}
						if(nightSta <= dt3 && dt3 < nightEnd)
						{
							listView1.Items[0].ForeColor = Color.Purple;//赤
						}
		            }

		            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
				}
			}
			else if(isHS)//HS
			{
				string shimeSign = "";
				string seikei1KataNo = "";
				string seikei1Tkeisu = "";
				string seikei1Zhosei = "";
				string seikei1Zvalue = "";
				string seikei1ct1up = "";
				string seikei1ct1dn = "";
				string seikei1cc2_cc1 = "";
				string seikei1cp = "";
				string seikei2cp = "";
				string seikei1TotalTime = "";
				string dateValue = "";
				string timeValue = "";
				string nikuUpLimit = "";
				string nikuLoLimit = "";
				string seikei1TounyuNo = "";
				string tounyuValue = "";

				if(!isMulti)//ｼﾝｸﾞﾙHS
				{
					string nikuData = "";
					string nikuResult = "";
					string resultCause = "";

		            ColumnHeader columnShime;
		            ColumnHeader columnSleeve;
		            ColumnHeader columnTkeisu;
		            ColumnHeader columnInitialTemp;
		            ColumnHeader columnSeikeiTemp;
		            ColumnHeader columnKaatsuTime;
		            ColumnHeader columnZ3Value;
		            ColumnHeader columnZ3hosei;
		            ColumnHeader columnCpValue;
		            ColumnHeader columnNikuatsuValue;
		            ColumnHeader columnNikuatsuUpper;
		            ColumnHeader columnNikuatsuLower;
		            ColumnHeader columnTact;
		            ColumnHeader columnDate;
		            ColumnHeader columnTimeStamp;
		            ColumnHeader columnResult;
		            ColumnHeader columnNgCause;
		            ColumnHeader columnShotCount;
		            ColumnHeader columnSeikeiCount;

					columnShime = new ColumnHeader();
		            columnSleeve = new ColumnHeader();
		            columnTkeisu = new ColumnHeader();
		            columnInitialTemp = new ColumnHeader();
		            columnSeikeiTemp = new ColumnHeader();
		            columnKaatsuTime = new ColumnHeader();
		            columnZ3Value = new ColumnHeader();
		            columnZ3hosei = new ColumnHeader();
		            columnCpValue = new ColumnHeader();
		            columnNikuatsuValue = new ColumnHeader();
		            columnNikuatsuUpper = new ColumnHeader();
		            columnNikuatsuLower = new ColumnHeader();
		            columnTact = new ColumnHeader();
		            columnTimeStamp = new ColumnHeader();
					columnDate = new ColumnHeader();
					columnResult = new ColumnHeader();
					columnNgCause = new ColumnHeader();
					columnShotCount = new ColumnHeader();
					columnSeikeiCount = new ColumnHeader();

					columnShime.Text = "締め";
		            columnSleeve.Text = "SL";
		            columnTkeisu.Text = "T係数";
		            columnInitialTemp.Text = "初期温度";
		            columnSeikeiTemp.Text = "成型温度";
		            columnKaatsuTime.Text = "加圧時間";
		            columnZ3Value.Text = "Z3";
		            columnZ3hosei.Text = "Z3補正";
		            columnCpValue.Text = "CP値";
		            columnTact.Text = "ﾀｸﾄ";
					columnDate.Text = "日付";
		            columnTimeStamp.Text = "時刻";
		            columnNikuatsuUpper.Text = "肉厚上限";
		            columnNikuatsuValue.Text = "測定値";
		            columnNikuatsuLower.Text = "肉厚下限";
		            columnResult.Text = "合否";
		            columnNgCause.Text = "原因";
		            columnShotCount.Text = "ｼｮｯﾄ数";
		            columnSeikeiCount.Text = "成型数";

		            ColumnHeader[] colHeaderRegValue =
		              { columnShime, columnSleeve, columnTkeisu, columnZ3hosei, columnZ3Value, columnInitialTemp, columnSeikeiTemp, columnKaatsuTime, 
		              columnCpValue, columnTact, columnDate, columnTimeStamp, columnNikuatsuUpper, columnNikuatsuValue, columnNikuatsuLower, columnResult, columnNgCause, columnShotCount, columnSeikeiCount};
		            listView1.Columns.AddRange(colHeaderRegValue);

					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);


		            var readToEnd = File.ReadAllLines(currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
		            int lines = readToEnd.Length;

		            for (int i = 0; i < lines; i++)
		            {
						if(i == 0)//ヘッダ部はスキップ
						{
							continue;
						}
		                //１行のstringをstream化してTextFieldParserで処理する
		                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
		                {
		                    using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
		                    {
		                        parser.TextFieldType = FieldType.Delimited;
		                        parser.Delimiters = new[] { "," };
		                        parser.HasFieldsEnclosedInQuotes = true;
		                        parser.TrimWhiteSpace = false;
		                        string[] fields = parser.ReadFields();

		                        for (int j = 0; j < fields.Length; j++)
		                        {
									if(j == 0)//date
									{
										dateValue = fields[j];
									}
									else if(j == 4)//ログファイル名
									{
									}
									else if(j == 6)//time
									{
										string datetimeData = fields[j];
										timeValue = datetimeData.Substring(datetimeData.Length - 8);
							            dateValue = datetimeData.Substring(0, 10);//日付
									}
									else if(j == 2)//seikeikiName
									{
										machineName = fields[j];
									}
									else if(j == 3)//seihinName
									{
									}
									else if(j == 4)//lastFileName
									{
									}
									else if(j == 5)//shimeSign
									{
										shimeSign = fields[j];
									}
									else if(j == 11)//sleeveNo
									{
										seikei1KataNo = fields[j];
									}
									else if(j == 10)//totalTime
									{
										seikei1TotalTime = fields[j];
									}
									else if(j == 17)//z3Value
									{
										seikei1Zvalue = fields[j];
									}
									else if(j == 15)//ct1Value
									{
										seikei1ct1up = fields[j];
									}
									else if(j == 16)//ct2Value
									{
										seikei1ct1dn = fields[j];
									}
									else if(j == 18)//cc32Value
									{
										seikei1cc2_cc1 = fields[j];
									}
									else if(j == 19)//cp1Value
									{
										seikei1cp = fields[j];
									}
									else if(j == 13)//Tkeisu
									{
										seikei1Tkeisu = fields[j];
									}
									else if(j == 14)//z3hosei
									{
										seikei1Zhosei = fields[j];
									}
									else if(j == 41)//cp2Value
									{
										seikei2cp = fields[j];
									}
									else if(j == 57)//nikuUpLimit
									{
										nikuUpLimit = fields[j];
									}
									else if(j == 58)//nikuData
									{
										nikuData = fields[j];
									}
									else if(j == 59)//nikuLoLimit
									{
										nikuLoLimit = fields[j];
									}
									else if(j == 60)//nikuResult
									{
										nikuResult = fields[j];
									}
									else if(j == 61)//resultCause
									{
										resultCause = fields[j];
									}
									else if(j == 62)//currentHousharitsu
									{
										ratioOfHousha = fields[j];
									}
									else if(j == 63)//currentOperator
									{
										opratorName = fields[j];
									}
									else if(j == 65)//shotCount
									{
	                                    seikei1TounyuNo = fields[j];
									}
									else if(j == 66)//seikeiCount
									{
	                                    tounyuValue = fields[j];
									}
		                        }
		                    }
		                    
		                }

						string[] item1 = {shimeSign, seikei1KataNo, seikei1Tkeisu, seikei1Zhosei, seikei1Zvalue, seikei1ct1up, seikei1ct1dn, seikei1cc2_cc1, 
						seikei2cp, seikei1TotalTime, dateValue, timeValue, 
						nikuUpLimit, nikuData, nikuLoLimit, nikuResult, resultCause, seikei1TounyuNo, tounyuValue};

						listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
			            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);

						if(nikuResult == "OK")
						{
				            listView1.Items[0].BackColor = Color.Lime;
						}
						else if(nikuResult == "NG")
						{
				            listView1.Items[0].BackColor = Color.Red;
						}

						if(seikei1KataNo == "0" || seikei1KataNo == "")
						{
							listView1.Items[0].BackColor = Color.Gray;
						}

						//日勤時間帯
						string strNoonSta = "08:00:00";
						string strNoonEnd = "17:00:00";
						DateTime noonSta = DateTime.Parse(strNoonSta);
						DateTime noonEnd = DateTime.Parse(strNoonEnd);

						//夕勤時間帯
						string strSunsetSta = "17:00:00";
						string strSunsetEnd = "23:59:59";
						DateTime sunsetSta = DateTime.Parse(strSunsetSta);
						DateTime sunsetEnd = DateTime.Parse(strSunsetEnd);

						//夜勤時間帯
						string strNightSta = "00:00:01";
						string strNightEnd = "08:00:00";
						DateTime nightSta = DateTime.Parse(strNightSta);
						DateTime nightEnd = DateTime.Parse(strNightEnd);

						DateTime dt3 = DateTime.Parse(timeValue);//タイムスタンプ(文字列)→DateTimeに変換

						if(noonSta <= dt3 && dt3 < noonEnd)
						{
							listView1.Items[0].ForeColor = Color.Blue;//青
						}
						if(sunsetSta <= dt3 && dt3 < sunsetEnd)
						{
							listView1.Items[0].ForeColor = Color.Green;//緑
						}
						if(nightSta <= dt3 && dt3 < nightEnd)
						{
							listView1.Items[0].ForeColor = Color.Purple;//赤
						}
		            }

		            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
				}
				else//ﾏﾙﾁHS
				{
					string [] nikuData = new string[] {"", "", "", "", "", ""};
					string [] nikuResult = new string[] {"", "", "", "", "", ""};
					string [] resultCause = new string[] {"", "", "", "", "", ""};

		            ColumnHeader columnShime;
		            ColumnHeader columnSleeve;
		            ColumnHeader columnTkeisu;
		            ColumnHeader columnInitialTemp;
		            ColumnHeader columnSeikeiTemp;
		            ColumnHeader columnKaatsuTime;
		            ColumnHeader columnZ3Value;
		            ColumnHeader columnZ3hosei;
		            ColumnHeader columnCpValue;
		            ColumnHeader columnNikuatsuValue1;
		            ColumnHeader columnNikuatsuValue2;
		            ColumnHeader columnNikuatsuValue3;
		            ColumnHeader columnNikuatsuValue4;
		            ColumnHeader columnNikuatsuValue5;
		            ColumnHeader columnNikuatsuValue6;
		            ColumnHeader columnResult1;
		            ColumnHeader columnResult2;
		            ColumnHeader columnResult3;
		            ColumnHeader columnResult4;
		            ColumnHeader columnResult5;
		            ColumnHeader columnResult6;
		            ColumnHeader columnNgCause1;
		            ColumnHeader columnNgCause2;
		            ColumnHeader columnNgCause3;
		            ColumnHeader columnNgCause4;
		            ColumnHeader columnNgCause5;
		            ColumnHeader columnNgCause6;
		            ColumnHeader columnNikuatsuUpper;
		            ColumnHeader columnNikuatsuLower;
		            ColumnHeader columnTact;
		            ColumnHeader columnDate;
		            ColumnHeader columnTimeStamp;
		            ColumnHeader columnShotCount;
		            ColumnHeader columnSeikeiCount;

		            columnShime = new ColumnHeader();
		            columnSleeve = new ColumnHeader();
		            columnTkeisu = new ColumnHeader();
		            columnInitialTemp = new ColumnHeader();
		            columnSeikeiTemp = new ColumnHeader();
		            columnKaatsuTime = new ColumnHeader();
		            columnZ3Value = new ColumnHeader();
		            columnZ3hosei = new ColumnHeader();
		            columnCpValue = new ColumnHeader();
		            columnNikuatsuValue1 = new ColumnHeader();
		            columnNikuatsuValue2 = new ColumnHeader();
		            columnNikuatsuValue3 = new ColumnHeader();
		            columnNikuatsuValue4 = new ColumnHeader();
		            columnNikuatsuValue5 = new ColumnHeader();
		            columnNikuatsuValue6 = new ColumnHeader();
		            columnNikuatsuUpper = new ColumnHeader();
		            columnNikuatsuLower = new ColumnHeader();
		            columnTact = new ColumnHeader();
		            columnTimeStamp = new ColumnHeader();
					columnDate = new ColumnHeader();
					columnResult1 = new ColumnHeader();
					columnResult2 = new ColumnHeader();
					columnResult3 = new ColumnHeader();
					columnResult4 = new ColumnHeader();
					columnResult5 = new ColumnHeader();
					columnResult6 = new ColumnHeader();
					columnNgCause1 = new ColumnHeader();
					columnNgCause2 = new ColumnHeader();
					columnNgCause3 = new ColumnHeader();
					columnNgCause4 = new ColumnHeader();
					columnNgCause5 = new ColumnHeader();
					columnNgCause6 = new ColumnHeader();
					columnShotCount = new ColumnHeader();
					columnSeikeiCount = new ColumnHeader();

					columnShime.Text = "締め";
		            columnSleeve.Text = "SL";
		            columnTkeisu.Text = "T係数";
		            columnInitialTemp.Text = "初期温度";
		            columnSeikeiTemp.Text = "成型温度";
		            columnKaatsuTime.Text = "加圧時間";
		            columnZ3Value.Text = "Z3";
		            columnZ3hosei.Text = "Z3補正";
		            columnCpValue.Text = "CP値";
		            columnTact.Text = "ﾀｸﾄ";
					columnDate.Text = "日付";
		            columnTimeStamp.Text = "時刻";
		            columnNikuatsuUpper.Text = "上限";
		            columnNikuatsuLower.Text = "下限";
		            columnNikuatsuValue1.Text = "Cav1";
		            columnResult1.Text = "Cav1合否";
		            columnNgCause1.Text = "Cav1原因";
		            columnNikuatsuValue2.Text = "Cav2";
		            columnResult2.Text = "Cav2合否";
		            columnNgCause2.Text = "Cav2原因";
		            columnNikuatsuValue3.Text = "Cav3";
		            columnResult3.Text = "Cav3合否";
		            columnNgCause3.Text = "Cav3原因";
		            columnNikuatsuValue4.Text = "Cav4";
		            columnResult4.Text = "Cav4合否";
		            columnNgCause4.Text = "Cav4原因";
		            columnNikuatsuValue5.Text = "Cav5";
		            columnResult5.Text = "Cav5合否";
		            columnNgCause5.Text = "Cav5原因";
		            columnNikuatsuValue6.Text = "Cav6";
		            columnResult6.Text = "Cav6合否";
		            columnNgCause6.Text = "Cav6原因";
		            columnShotCount.Text = "ｼｮｯﾄ数";
		            columnSeikeiCount.Text = "成型数";

		            ColumnHeader[] colHeaderRegValue =
		              { columnShime, columnSleeve, columnTkeisu, columnZ3hosei, columnZ3Value, columnInitialTemp, columnSeikeiTemp, columnKaatsuTime, 
		              columnCpValue, columnTact, columnDate, columnTimeStamp, columnNikuatsuUpper, columnNikuatsuLower, 
		              columnNikuatsuValue1, columnResult1, columnNgCause1, 
		              columnNikuatsuValue2, columnResult2, columnNgCause2, 
		              columnNikuatsuValue3, columnResult3, columnNgCause3, 
		              columnNikuatsuValue4, columnResult4, columnNgCause4, 
		              columnNikuatsuValue5, columnResult5, columnNgCause5, 
		              columnNikuatsuValue6, columnResult6, columnNgCause6, 
		              columnShotCount, columnSeikeiCount};
		            listView1.Columns.AddRange(colHeaderRegValue);

					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

		            var readToEnd = File.ReadAllLines(currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
		            int lines = readToEnd.Length;

		            for (int i = 0; i < lines; i++)
		            {
						if(i == 0)//ヘッダ部はスキップ
						{
							continue;
						}
		                //１行のstringをstream化してTextFieldParserで処理する
		                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
		                {
		                    using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
		                    {
		                        parser.TextFieldType = FieldType.Delimited;
		                        parser.Delimiters = new[] { "," };
		                        parser.HasFieldsEnclosedInQuotes = true;
		                        parser.TrimWhiteSpace = false;
		                        string[] fields = parser.ReadFields();

		                        for (int j = 0; j < fields.Length; j++)
		                        {
									if(j == 0)//date
									{
										dateValue = fields[j];
									}
									else if(j == 4)//ログファイル名
									{
										if(i == (lines - 1))//最新のログ位置
										{
										}
									}
									else if(j == 6)//time
									{
										string datetimeData = fields[j];
										timeValue = datetimeData.Substring(datetimeData.Length - 8);
							            dateValue = datetimeData.Substring(0, 10);//日付
									}
									else if(j == 2)//seikeikiName
									{
										machineName = fields[j];
									}
									else if(j == 3)//seihinName
									{
									}
									else if(j == 4)//lastFileName
									{
									}
									else if(j == 5)//shimeSign
									{
										shimeSign = fields[j];
									}
									else if(j == 11)//sleeveNo
									{
										seikei1KataNo = fields[j];
									}
									else if(j == 10)//totalTime
									{
										seikei1TotalTime = fields[j];
									}
									else if(j == 17)//z3Value
									{
										seikei1Zvalue = fields[j];
									}
									else if(j == 15)//ct1Value
									{
										seikei1ct1up = fields[j];
									}
									else if(j == 16)//ct2Value
									{
										seikei1ct1dn = fields[j];
									}
									else if(j == 18)//cc32Value
									{
										seikei1cc2_cc1 = fields[j];
									}
									else if(j == 19)//cp1Value
									{
										seikei1cp = fields[j];
									}
									else if(j == 13)//Tkeisu
									{
										seikei1Tkeisu = fields[j];
									}
									else if(j == 14)//z3hosei
									{
										seikei1Zhosei = fields[j];
									}
									else if(j == 41)//cp2Value
									{
										seikei2cp = fields[j];
									}
									else if(j == 57)//nikuUpLimit
									{
										nikuUpLimit = fields[j];
									}
									else if(j == 58)//nikuLoLimit
									{
										nikuLoLimit = fields[j];
									}
									else if(j == 59)//nikuData(Cav1)
									{
										nikuData[0] = fields[j];
									}
									else if(j == 62)//nikuData(Cav2)
									{
										nikuData[1] = fields[j];
									}
									else if(j == 65)//nikuData(Cav3)
									{
										nikuData[2] = fields[j];
									}
									else if(j == 68)//nikuData(Cav4)
									{
										nikuData[3] = fields[j];
									}
									else if(j == 71)//nikuData(Cav5)
									{
										nikuData[4] = fields[j];
									}
									else if(j == 74)//nikuData(Cav6)
									{
										nikuData[5] = fields[j];
									}

									else if(j == 60)//nikuResult(Cav1)
									{
										nikuResult[0] = fields[j];
									}
									else if(j == 63)//nikuResult(Cav2)
									{
										nikuResult[1] = fields[j];
									}
									else if(j == 66)//nikuResult(Cav3)
									{
										nikuResult[2] = fields[j];
									}
									else if(j == 69)//nikuResult(Cav4)
									{
										nikuResult[3] = fields[j];
									}
									else if(j == 72)//nikuResult(Cav5)
									{
										nikuResult[4] = fields[j];
									}
									else if(j == 75)//nikuResult(Cav6)
									{
										nikuResult[5] = fields[j];
									}

									else if(j == 61)//resultCause(Cav1)
									{
										resultCause[0] = fields[j];
									}
									else if(j == 64)//resultCause(Cav2)
									{
										resultCause[1] = fields[j];
									}
									else if(j == 67)//resultCause(Cav3)
									{
										resultCause[2] = fields[j];
									}
									else if(j == 70)//resultCause(Cav4)
									{
										resultCause[3] = fields[j];
									}
									else if(j == 73)//resultCause(Cav5)
									{
										resultCause[4] = fields[j];
									}
									else if(j == 76)//resultCause(Cav6)
									{
										resultCause[5] = fields[j];
									}

									else if(j == 77)//currentHousharitsu
									{
										ratioOfHousha = fields[j];
									}
									else if(j == 78)//currentOperator
									{
										opratorName = fields[j];
									}
									else if(j == 80)//shotCount
									{
	                                    seikei1TounyuNo = fields[j];
									}
									else if(j == 81)//seikeiCount
									{
	                                    tounyuValue = fields[j];
									}
		                        }
		                    }
		                    
		                }

						string[] item1 = {shimeSign, seikei1KataNo, seikei1Tkeisu, seikei1Zhosei, seikei1Zvalue, seikei1ct1up, seikei1ct1dn, seikei1cc2_cc1, 
						seikei2cp, seikei1TotalTime, dateValue, timeValue, 
						nikuUpLimit, nikuLoLimit, 
						nikuData[0], nikuResult[0], resultCause[0], 
						nikuData[1], nikuResult[1], resultCause[1], 
						nikuData[2], nikuResult[2], resultCause[2], 
						nikuData[3], nikuResult[3], resultCause[3], 
						nikuData[4], nikuResult[4], resultCause[4], 
						nikuData[5], nikuResult[5], resultCause[5], 
						seikei1TounyuNo, tounyuValue};

						listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
			            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);

						if(seikei1KataNo == "0" || seikei1KataNo == "")
						{
							listView1.Items[0].BackColor = Color.Gray;
						}

						//日勤時間帯
						string strNoonSta = "08:00:00";
						string strNoonEnd = "17:00:00";
						DateTime noonSta = DateTime.Parse(strNoonSta);
						DateTime noonEnd = DateTime.Parse(strNoonEnd);

						//夕勤時間帯
						string strSunsetSta = "17:00:00";
						string strSunsetEnd = "23:59:59";
						DateTime sunsetSta = DateTime.Parse(strSunsetSta);
						DateTime sunsetEnd = DateTime.Parse(strSunsetEnd);

						//夜勤時間帯
						string strNightSta = "00:00:01";
						string strNightEnd = "08:00:00";
						DateTime nightSta = DateTime.Parse(strNightSta);
						DateTime nightEnd = DateTime.Parse(strNightEnd);

						DateTime dt3 = DateTime.Parse(timeValue);//タイムスタンプ(文字列)→DateTimeに変換

						if(noonSta <= dt3 && dt3 < noonEnd)
						{
							listView1.Items[0].ForeColor = Color.Blue;//青
						}
						if(sunsetSta <= dt3 && dt3 < sunsetEnd)
						{
							listView1.Items[0].ForeColor = Color.Green;//緑
						}
						if(nightSta <= dt3 && dt3 < nightEnd)
						{
							listView1.Items[0].ForeColor = Color.Purple;//赤
						}
		            }

		            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

				}
			}



        }

    }
}
