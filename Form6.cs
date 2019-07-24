using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Collections;
using Microsoft.VisualBasic.FileIO;

namespace PgmDataMonitor
{
    public partial class Form6 : Form
    {
        public struct SleeveList
        {
            public string sleeveNumber;
            public int shotCount;
			public int seikeiSuu;
			public int ryohinSuu;
			public int furyouSuu;
			public int nikuatsuNG;//肉厚不良
			public int kizuNG;//キズ
			public int butsuNG;//ブツ
			public int yakeNG;//ヤケ
			public int hibicrackNG;//ヒビ/クラック
			public int gasukizuNG;//ガスキズ
			public int houshakizuNG;//放射キズ
			public int giratsukikumoriNG;//ギラツキ/クモリ
			public int hennikumendareNG;//偏肉/面ダレ
			public int yogoreNG;//汚れ
			public int hokoriNG;//ほこり
			public int keijoseidoNG;//形状精度
			public int etcNG;//その他
        }

		public struct sameKataList
		{
			public string sleeveNumber;
			public string [] result;
			public int lineNumber;
			public bool [] isEdit;
		}

        static public SYSSET SETDATA = new SYSSET();

		public int totalSleeve = 0;
		public int currentSleeve = 0;
		public int seikeiCount;
		public int updateMode = 0;

		public string currentOperator = "";
		public string currentHousharitsu = "";

		public string shimeSign = "";
		public string ct1Value = "";
		public string ct2Value = "";
		public string timeValue = "";

		public string nikuEdit = "";
		public string causeEdit = "";
		public string nikuUpLimit = "";
		public string nikuLoLimit = "";
		public string [] nikuData = new string[] {"", "", "", "", "", ""};
		public string [] nikuResult = new string[] {"", "", "", "", "", ""};
		public string [] resultCause = new string[] {"", "", "", "", "", ""};

		public string backCsvFile = "";
		public string backTimeValue = "";

		public string dateValue = "";
        public string datetimeValue;
		public string hinshuValue;
		public string tounyuValue;
		public string samplingValue;
		public string processValue;
//成型1
		public string seikei1KataNo;
		public string seikei1TounyuNo;
		public string seikei1Tkeisu;
		public string seikei1Zhosei;
		public string seikei1ct1up;
		public string seikei1ct1dn;
		public string seikei1Zvalue;
		public string seikei1cc2_cc1;
		public string seikei1cp;
		public string seikei1Step1Time;
		public string seikei1Step2Time;
		public string seikei1Step3Time;
		public string seikei1Step4Time;
		public string seikei1Step5Time;
		public string seikei1Step6Time;
		public string seikei1Step7Time;
		public string seikei1Step8Time;
		public string seikei1Step9Time;
		public string seikei1Step10Time;
		public string seikei1Step11Time;
		public string seikei1Step12Time;

		public string seikei1TotalTime;
//成型2
		public string seikei2KataNo;
		public string seikei2TounyuNo;
		public string seikei2Tkeisu;
		public string seikei2Zhosei;
		public string seikei2ct1up;
		public string seikei2ct1dn;
		public string seikei2Zvalue;
		public string seikei2cc2_cc1;
		public string seikei2cp;
		public string seikei2Step1Time;
		public string seikei2Step2Time;
		public string seikei2Step3Time;
		public string seikei2Step4Time;
		public string seikei2Step5Time;
		public string seikei2Step6Time;
		public string seikei2Step7Time;
		public string seikei2Step8Time;
		public string seikei2Step9Time;
		public string seikei2Step10Time;
		public string seikei2Step11Time;
		public string seikei2Step12Time;

		public string seikei2TotalTime;

		public string shukaiCount;
		public string errorStr;

		string[] headerValues;
		
		public DateTime backFileTime;
		public string currentCsvFile = "";
		public StreamWriter logWriter;
		
		public string currentSeihin = "";
		public string currentMachine = "";
		public string backSeihinName = "";

		public string currentKaatsujikan = "";
		public int currentKaatsujikanUp = 0;
		public int currentKaatsujikanLo = 0;

        public int hidden = 1;

		FormCalender formCalender = null;
		SettingForm setform = null;
		public string lastFileName = "";

        public Form6()
        {
            InitializeComponent();
			ReadDataFromXml();

			for(int i = 0; i < SETDATA.OperatorName.Length; i++)//作業者名
			{
				comboBox12.Items.Add(SETDATA.OperatorName[i]);
			}

#if false//書き込み用 start
			WriteDataToXml();
#endif //書き込み用 end

            // ListViewコントロールのプロパティを設定
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.Sorting = SortOrder.Ascending;
            listView1.View = View.Details;
            listView1.HideSelection = false;
//            listView1.LabelEdit = true;


            // 列（コラム）ヘッダの作成
            ColumnHeader columnShime;
            ColumnHeader columnSleeve;
            ColumnHeader columnTkeisu;
            ColumnHeader columnInitialTemp;
            ColumnHeader columnSeikeiTemp;
            ColumnHeader columnKaatsuTime;
            ColumnHeader columnZ3Value;
            ColumnHeader columnZ3hosei;
//            ColumnHeader columnCC1Value;
//            ColumnHeader columnCC2Value;
//            ColumnHeader columnCC3Value;
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
//            columnCC1Value = new ColumnHeader();
//            columnCC2Value = new ColumnHeader();
//            columnCC3Value = new ColumnHeader();
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

			columnShime.Text = "締め位置";
            columnSleeve.Text = "SL";
            columnTkeisu.Text = "T係数";
            columnInitialTemp.Text = "初期温度";
            columnSeikeiTemp.Text = "成型温度";
            columnKaatsuTime.Text = "加圧時間";
            columnZ3Value.Text = "Z3";
            columnZ3hosei.Text = "Z3補正";
//            columnCC1Value.Text = "CC1";
//            columnCC2Value.Text = "CC2";
//            columnCC3Value.Text = "CC3";
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
              { columnShime, columnSleeve, columnTkeisu, columnInitialTemp, columnSeikeiTemp, columnKaatsuTime, columnZ3Value, columnZ3hosei, /*columnCC1Value, columnCC2Value, columnCC3Value, */
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

			listView1.HideSelection = false;

			string searchStr = "";
			if(SETDATA.machineType == 5)
			{
				searchStr = "HS";
			}
			for(int i = 0; i < SETDATA.goukiName.Length; i++)//成型機No
			{
				if(SETDATA.goukiName[i].IndexOf(searchStr) != -1)
				{
					comboBox2.Items.Add(SETDATA.goukiName[i]);
				}
			}
			for(int i = 0; i < Form1.SETDATA.ngCause.Length; i++)//不良原因
			{
				comboBox3.Items.Add(Form1.SETDATA.ngCause[i]);
			}

			this.Width = SETDATA.windowWidth;
			this.Height = SETDATA.windowHeight;
			listView1.Width = SETDATA.listviewWidth;
			listView1.Height = SETDATA.listviewHeight;

			System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
			System.Version ver = asm.GetName().Version;
            this.Text += "  Ver:" + ver;

            this.AutoScroll = true;

			//フォルダにある最新のCSVを検索
			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(stCurrentDir);

			string existFile = "";
			if(SETDATA.machineType == 5)
			{
				existFile = "*_HS*号機_マルチCav.csv";
			}
			System.IO.FileInfo[] files = di.GetFiles(existFile, System.IO.SearchOption.TopDirectoryOnly);

            string strTime = "2019/01/01 00:00:00";
			DateTime lastdt = DateTime.Parse(strTime);
			DateTime tempdt;
			int lastIndex = 0xFFFF;
			for(int i = 0; i < files.Length; i++)
			{
				tempdt = files[i].LastWriteTime;
				if(lastdt < tempdt)
				{
					lastdt = tempdt;
					lastIndex = i;
				}
			}

			if(lastIndex != 0xFFFF)
			{
				//全体の色設定
	            listView1.Sorting = SortOrder.None;
	            listView1.ForeColor = Color.Black;//初期の色
	            listView1.BackColor = Color.Pink;//全体背景色

	            //CSV読込。高速化を狙い、最大行数を取得後、for分でループする
	            currentCsvFile = files[lastIndex].FullName;
	            var readToEnd = File.ReadAllLines(currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
	            int lines = readToEnd.Length;

				string machineName = "";
				string codeName = "";
				string ratioOfHousha = "";
				string opratorName = "";

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
										backCsvFile = fields[j];
									}
								}
								else if(j == 6)//time
								{
									string datetimeData = fields[j];
									timeValue = datetimeData.Substring(datetimeData.Length - 8);
								}
								else if(j == 2)//seikeikiName
								{
									machineName = fields[j];
								}
								else if(j == 3)//seihinName
								{
									backSeihinName = codeName = fields[j];
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
									ct1Value = fields[j];
								}
								else if(j == 16)//ct2Value
								{
									ct2Value = fields[j];
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
								else if(j == 79)//shotCount
								{
                                    seikei1TounyuNo = fields[j];
								}
								else if(j == 80)//seikeiCount
								{
                                    tounyuValue = fields[j];
								}
	                        }
	                    }
	                    
	                }

					string[] item1 = {shimeSign, seikei1KataNo, seikei1Tkeisu, seikei1ct1up, seikei1ct1dn, seikei1cc2_cc1, seikei1Zvalue, seikei1Zhosei, 
					/*cc1Value, cc2Value, cc3Value,*/ seikei2cp, seikei1TotalTime, dateValue, timeValue, 
					nikuUpLimit, nikuLoLimit, 
					nikuData[0], nikuResult[0], resultCause[0], 
					nikuData[1], nikuResult[1], resultCause[1], 
					nikuData[2], nikuResult[2], resultCause[2], 
					nikuData[3], nikuResult[3], resultCause[3], 
					nikuData[4], nikuResult[4], resultCause[4], 
					nikuData[5], nikuResult[5], resultCause[5], 
					seikei1TounyuNo, tounyuValue};

					listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
		            listView1.Font = new System.Drawing.Font("Times New Roman", 10, System.Drawing.FontStyle.Regular);

					listView1.Items[0].UseItemStyleForSubItems = false;
					int jj = 0;
					for(int k = 0; k < nikuResult.Length; k++)
					{
						listView1.Items[0].SubItems[14 + jj].Text = nikuData[k];
						listView1.Items[0].SubItems[15 + jj].Text = nikuResult[k];
						listView1.Items[0].SubItems[16 + jj].Text = resultCause[k];
						if(nikuResult[k] == "-")
						{
							listView1.Items[0].SubItems[14 + jj].BackColor = Color.White;
							listView1.Items[0].SubItems[15 + jj].BackColor = Color.White;
							listView1.Items[0].SubItems[16 + jj].BackColor = Color.White;
						}
						else
						{
							if(nikuResult[k] == "OK")
							{
								listView1.Items[0].SubItems[14 + jj].BackColor = Color.Lime;
								listView1.Items[0].SubItems[15 + jj].BackColor = Color.Lime;
								listView1.Items[0].SubItems[16 + jj].BackColor = Color.Lime;
							}
							else if(nikuResult[k] == "NG")
							{
								listView1.Items[0].SubItems[14 + jj].BackColor = Color.Red;
								listView1.Items[0].SubItems[15 + jj].BackColor = Color.Red;
								listView1.Items[0].SubItems[16 + jj].BackColor = Color.Red;
							}
						}
						jj += 3;
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

	            listView1.Font = new System.Drawing.Font("Times New Roman", 14, System.Drawing.FontStyle.Regular);
				//ヘッダの幅を自動調節
				listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

				comboBox2.Text = currentMachine = machineName;
				label1.Text = currentSeihin = codeName;
				numericUpDown12.Text = nikuUpLimit;
				numericUpDown13.Text = nikuLoLimit;
				numericUpDown2.Text = ratioOfHousha;
				comboBox12.Text = opratorName;

				//加圧時間の決定
				currentKaatsujikan = codeName;
				for(int i = 0; i < SETDATA.seihinName.Length; i++)
				{
					if(SETDATA.seihinName[i] == currentKaatsujikan)
					{
						if(SETDATA.KaatsuJikanUpper[i] != "" && SETDATA.KaatsuJikanLower[i] != "")
						{
							currentKaatsujikanUp = int.Parse(SETDATA.KaatsuJikanUpper[i]);
							currentKaatsujikanLo = int.Parse(SETDATA.KaatsuJikanLower[i]);
							break;
						}
					}
				}

				//加圧時間の範囲外なら、強調表示
				for(int i = 0; i < listView1.Items.Count; i++)
				{
					string tmp_cc32Value = listView1.Items[i].SubItems[5].Text;
					double kaatsuTime = double.Parse(tmp_cc32Value);

					string slNum = listView1.Items[i].SubItems[1].Text;
					if(slNum != "0" && slNum != "")
					{
						if(kaatsuTime < currentKaatsujikanLo || currentKaatsujikanUp < kaatsuTime)
						{
							listView1.Items[i].ForeColor = Color.Yellow;
				            listView1.Items[i].Font = new System.Drawing.Font("Times New Roman", 14, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic);
						}
					}
				}
				//ヘッダの幅を自動調節
				listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

				numericUpDown3.Text = currentKaatsujikanUp.ToString();
				numericUpDown4.Text = currentKaatsujikanLo.ToString();

				timer1.Enabled = true;
				button1.Enabled = false;
				return;
			}
			//直前の成型号機
			comboBox2.Text = SETDATA.selectedMachine;
			//直前の作業者
			comboBox12.Text = SETDATA.selectedOperator;

			if(SETDATA.selectedMachine != "" && SETDATA.selectedOperator != "")
			{
				timer1.Enabled = true;
				button1.Enabled = false;
			}
			else
			{
				button1.Enabled = true;
			}
        }

        private void ParseLogString(string [] strResults)
		{
			datetimeValue = strResults[0];
			timeValue = datetimeValue.Substring(datetimeValue.Length - 8);

			hinshuValue = strResults[1];

            tounyuValue = strResults[2];
			samplingValue = strResults[3];
			processValue = strResults[4];
//成型1
			seikei1KataNo = strResults[5];
			seikei1TounyuNo = strResults[6];
			seikei1Tkeisu = strResults[7];
			string tmpTkeisu = string.Format("{0:#.###}", seikei1Tkeisu);
			seikei1Tkeisu = tmpTkeisu;

			seikei1Zhosei = strResults[8];
			string tmpZhosei = string.Format("{0:#.###}", seikei1Zhosei);
			seikei1Zhosei = tmpZhosei;

			seikei1ct1up = strResults[9];
			seikei1ct1dn = strResults[10];
			seikei1Zvalue = strResults[11];
			seikei1cc2_cc1 = strResults[12];
			seikei1cp = strResults[13];
			seikei1Step1Time = strResults[14];
			seikei1Step2Time = strResults[15];
			seikei1Step3Time = strResults[16];
			seikei1Step4Time = strResults[17];
			seikei1Step5Time = strResults[18];
			seikei1Step6Time = strResults[19];
			seikei1Step7Time = strResults[20];
			seikei1Step8Time = strResults[21];
			seikei1Step9Time = strResults[22];
			seikei1Step10Time = strResults[23];
			seikei1Step11Time = strResults[24];
			seikei1Step12Time = strResults[25];

			seikei1TotalTime = strResults[26];
//成型2
			seikei2KataNo = strResults[27];
			seikei2TounyuNo = strResults[28];
			seikei2Tkeisu = strResults[29];
			string tmp2Tkeisu = string.Format("{0:#.###}", seikei2Tkeisu);
			seikei2Tkeisu = tmp2Tkeisu;

			seikei2Zhosei = strResults[30];
			string tmp2Zhosei = string.Format("{0:#.###}", seikei2Zhosei);
			seikei2Zhosei = tmp2Zhosei;

			seikei2ct1up = strResults[31];
			seikei2ct1dn = strResults[32];
			seikei2Zvalue = strResults[33];
			seikei2cc2_cc1 = strResults[34];
			seikei2cp = strResults[35];
			seikei2Step1Time = strResults[36];
			seikei2Step2Time = strResults[37];
			seikei2Step3Time = strResults[38];
			seikei2Step4Time = strResults[39];
			seikei2Step5Time = strResults[40];
			seikei2Step6Time = strResults[41];
			seikei2Step7Time = strResults[42];
			seikei2Step8Time = strResults[43];
			seikei2Step9Time = strResults[44];
			seikei2Step10Time = strResults[45];
			seikei2Step11Time = strResults[46];
			seikei2Step12Time = strResults[47];

			seikei2TotalTime = strResults[48];

			shukaiCount = strResults[49];

            if(strResults.Length == 51)
            {
				errorStr = strResults[50];
			}

			shimeSign = "";
            DateTime dt = DateTime.Now;
//            dateValue = string.Format("{0}", dt.ToString("yyyy/MM/dd"));//日付
            dateValue = datetimeValue.Substring(0, 10);//日付

			string seikeikiName = "";
			string seikeiki = comboBox2.Text;
			if(seikeiki == "")//何も選択されていない場合
			{
				seikeikiName = "seikeiki";
			}
			else
			{
				seikeikiName = comboBox2.Text;
			}

			string logStr = "";
			logStr += "," + seikeikiName;
			logStr += "," + currentSeihin;
			logStr += "," + lastFileName;

			logStr += "," + shimeSign;

			logStr += "," + datetimeValue;
			logStr += "," + hinshuValue;
			logStr += "," + tounyuValue;
			logStr += "," + samplingValue;
			logStr += "," + processValue;

			logStr += "," + seikei1KataNo;
			logStr += "," + seikei1TounyuNo;
			logStr += "," + seikei1Tkeisu;
			logStr += "," + seikei1Zhosei;
			logStr += "," + seikei1ct1up;
			logStr += "," + seikei1ct1dn;
			logStr += "," + seikei1Zvalue;
			logStr += "," + seikei1cc2_cc1;
			logStr += "," + seikei1cp;
			logStr += "," + seikei1Step1Time;
			logStr += "," + seikei1Step2Time;
			logStr += "," + seikei1Step3Time;
			logStr += "," + seikei1Step4Time;
			logStr += "," + seikei1Step5Time;
			logStr += "," + seikei1Step6Time;
			logStr += "," + seikei1Step7Time;
			logStr += "," + seikei1Step8Time;
			logStr += "," + seikei1Step9Time;
			logStr += "," + seikei1Step10Time;
			logStr += "," + seikei1Step11Time;
			logStr += "," + seikei1Step12Time;

			logStr += "," + seikei1TotalTime;

			logStr += "," + seikei2KataNo;
			logStr += "," + seikei2TounyuNo;
			logStr += "," + seikei2Tkeisu;
			logStr += "," + seikei2Zhosei;
			logStr += "," + seikei2ct1up;
			logStr += "," + seikei2ct1dn;
			logStr += "," + seikei2Zvalue;
			logStr += "," + seikei2cc2_cc1;
			logStr += "," + seikei2cp;
			logStr += "," + seikei2Step1Time;
			logStr += "," + seikei2Step2Time;
			logStr += "," + seikei2Step3Time;
			logStr += "," + seikei2Step4Time;
			logStr += "," + seikei2Step5Time;
			logStr += "," + seikei2Step6Time;
			logStr += "," + seikei2Step7Time;
			logStr += "," + seikei2Step8Time;
			logStr += "," + seikei2Step9Time;
			logStr += "," + seikei2Step10Time;
			logStr += "," + seikei2Step11Time;
			logStr += "," + seikei2Step12Time;

			logStr += "," + seikei2TotalTime;

			logStr += "," + shukaiCount;
			logStr += "," + errorStr;

			//CSVに肉厚、上限、下限、合否、不良原因も設定
			logStr += string.Format(",{0:F3},{1:F3}",nikuUpLimit, nikuLoLimit);

			for(int i = 0; i < nikuData.Length; i++)
			{
				nikuData[i] = "";
				nikuResult[i] = "";
				resultCause[i] = "";
				
				logStr += string.Format(",{0:F3}",nikuData[i]);
				logStr += "," + nikuResult[i] + "," + resultCause[i];
			}

			currentOperator = comboBox12.Text;
			currentHousharitsu = numericUpDown2.Text;
			logStr += "," + currentHousharitsu + "," + currentOperator + "," + seikei1KataNo + "," + seikei1TounyuNo + "," + tounyuValue;

			if(backTimeValue != timeValue)//ログファイルが更新されていても、最新時間が同じ場合の念の為ガード
			{
			//csvファイルに更新がある場合
				string fileName = lastFileName.Substring(0, (lastFileName.Length - 4));
				//CSVに保存する
				if(WriteDataToCsv(logStr, currentSeihin, comboBox2.Text, headerValues))
				{
					backTimeValue = timeValue;

					//リストビューを更新
					AddListView();
				}
			}
		}

        private void timer1_Tick(object sender, EventArgs e)
        {
			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();


			int boundPos = stCurrentDir.LastIndexOf("\\");
            string str1 = stCurrentDir.Substring(0, boundPos);

			if(comboBox2.SelectedIndex == 0)//HS1のみ
			{
            	str1 += "\\Log";
			}
			else
			{
            	str1 += "\\Log\\result";
			}

            stCurrentDir = str1;

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(stCurrentDir);
            IEnumerable<System.IO.FileInfo> files = di.EnumerateFiles("*.csv", System.IO.SearchOption.TopDirectoryOnly);

			string strTime = "2019/01/01 01:23:45";
			DateTime lastDt = DateTime.Parse(strTime);
			DateTime destDt;

			SortedList sl = new SortedList();//ログファイル名と更新時間のペア
			
			//最近更新されたファイルの探索
			try
			{
	            foreach(System.IO.FileInfo f in files)
	            {
                    string pathStr = stCurrentDir + "\\" + f.Name;
                    destDt = System.IO.File.GetLastWriteTime(pathStr);
					if(lastDt < destDt)
					{
						lastDt = destDt;
						lastFileName = f.Name;
					}

					//ログファイルのリストに入れる(＆昇順に並べ替えている)
					sl.Add(destDt, f.Name);
	            }
	        }
			catch (System.IO.IOException ex)
			{
				string errorStr = "他のアプリがCSVファイルを開いている可能性があります";
			    System.Console.WriteLine(errorStr);
			    System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			    return;
			}

			string str_back = backFileTime.ToString();
			string str_last = lastDt.ToString();
			if(backCsvFile == lastFileName && str_back == str_last)//最新ファイル名が同一かつ更新日時が同一の場合
			{
				updateMode = 0;
				return;//更新なしと判断
			}

			//昇順に並べ変えたログファイルを文字列配列に入れる
            int aa = sl.Count;
            string[] fileList = new string[sl.Count];
            DateTime[] dateList = new DateTime[sl.Count];
            int ll = 0;
            foreach (object item in sl.Values)
            {
                fileList[ll] = item.ToString();
                string str = sl.GetKey(ll).ToString();
                dateList[ll] = DateTime.Parse(str);
                ll++;
            }

			//CSVファイルにある最新のログファイル名の最新のログ以降のログを読み込む。その後、さらに新しいログファイルがあれば全て読み込む。その為にループする
			bool isExist = false;//CSVファイルにあるログファイル名が現れたらfalse
			for(int k = 0; k < sl.Count; k++)
			{
				if(backCsvFile != "")
				{
					lastFileName = fileList[k];
					lastDt = dateList[k];
					if(!isExist && lastFileName != backCsvFile)
					{
						continue;
					}
					isExist = true;
				}
				else
				{
					k = sl.Count - 1;//次のループを終了させる
				}

	            string srcFile = stCurrentDir + "\\" + lastFileName;
	            string dstFile = stCurrentDir + "\\" + lastFileName + ".tmp";

				//tmpファイルがあれば、全て削除しておく
	            System.IO.DirectoryInfo tmp_di = new System.IO.DirectoryInfo(stCurrentDir);
	            IEnumerable<System.IO.FileInfo> files_tmp = tmp_di.EnumerateFiles("*.tmp", System.IO.SearchOption.TopDirectoryOnly);

				try
				{
		            foreach(System.IO.FileInfo f in files_tmp)
		            {
	                    string deleteFile = stCurrentDir + "\\" + f.Name;
	                    File.SetAttributes(deleteFile, FileAttributes.Normal);
						File.Delete(deleteFile);
		            }
		        }
				catch (System.IO.IOException ex)
				{
					string errorStr = "他のアプリがtmpファイルを開いている可能性があります";
				    System.Console.WriteLine(errorStr);
				    System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
				    continue;
				}

	            //ファイルを開く
				FileStream fp;
				StreamReader sr = null;
				try
				{
					//最新のCSVファイルをWORK用一時ファイルにコピーし、以降は一時ファイルを扱う。
					//最新CSVファイルは成型機VBアプリもアクセスするため、競合区間を最小限にするため。
					File.Copy(@srcFile, @dstFile);
		            fp = new FileStream(dstFile, FileMode.Open, FileAccess.Read);
		            sr = new StreamReader(fp, System.Text.Encoding.GetEncoding("Shift_JIS"));
				}
				catch (System.IO.IOException ex)
				{
					string errorStr = "CSVファイルコピー失敗またはCSVファイルが開けなかった可能性があります";
				    System.Console.WriteLine(errorStr);
				    System.Console.WriteLine(ex.Message);
					
					if(sr != null)
					{
						sr.Close();
					}
					File.Delete(@dstFile);//一時ファイルは削除
					LogFileOut(errorStr);
				    continue;
				}

				int pos = 0;
				int justPos = 0;
	            string line = "";
	            while (sr.EndOfStream == false)
	            {
	                line = sr.ReadLine();

					//先頭行のParse
					if(pos == 0)
					{	
						headerValues = line.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
	                }

	                pos++;
				}
				justPos = pos;

	            //最後行のParse
	            string lastline = line;
	            string[] strResults = lastline.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

	            if(strResults.Length <= 4)//異常系：成型条件変更等のログの場合、抜ける
	            {
					if(sr != null)
					{
						sr.Close();
					}
					File.Delete(@dstFile);//一時ファイルは削除
					continue;
				}

				//最新ファイル名が直前のファイル名と異なる場合、ファイル名と更新時間を入れ替え
				if(backCsvFile != lastFileName)
				{
					backCsvFile = lastFileName;
					backFileTime = lastDt;
					updateMode = 2;
				}
				//最新ファイル名が直前のファイル名と同じで時間が異なる場合、更新時間を入れ替え
				str_back = backFileTime.ToString();
				str_last = lastDt.ToString();
				if(backCsvFile == lastFileName && str_back != str_last)
				{
					backFileTime = lastDt;
					updateMode = 1;
				}

				hinshuValue = strResults[1];

	            currentSeihin = hinshuValue.Substring(0, 7);
				int noData = 0;
				for(int i = 0; i < SETDATA.seihinName.Length; i++)//製品名
				{
					if(SETDATA.seihinName[i] == currentSeihin)//製品に応じた肉厚上限、下限を設定
					{
						nikuUpLimit = SETDATA.nikuUpper[i];
						nikuLoLimit = SETDATA.nikuLower[i];
						double upper = double.Parse(nikuUpLimit);
						double lower = double.Parse(nikuLoLimit);
		
						double data = (upper + lower) / 2;

			            numericUpDown12.DecimalPlaces = 3;
			            numericUpDown13.DecimalPlaces = 3;
			            numericUpDown12.Value = (decimal)upper;
			            numericUpDown13.Value = (decimal)lower;
			            label1.Text = currentSeihin;

						break;
					}
					noData++;
				}

				if(noData == SETDATA.seihinName.Length)//製品登録が無い場合
				{
					nikuUpLimit = "1.222";
					nikuLoLimit = "1.111";
		            numericUpDown12.Text = "1.222";
		            numericUpDown13.Text = "1.111";
		            label1.Text = "製品登録して下さい";
				}

				//更新部分を検索する
				int lastIndex = 0;
				string lastTime = "";
				int lastPos = 0;
				if(updateMode == 1)//ログファイルが同一で、時間のみ更新があった場合
				{
					if(listView1.Items.Count > 0)//ListViewに1つでも登録がある
					{
						lastIndex = listView1.Items.Count;
						lastTime = listView1.Items[0].SubItems[11].Text;//最新のログの時間

						sr.DiscardBufferedData();//一度先頭に戻す
						sr.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
			            while (sr.EndOfStream == false)
			            {
			                line = sr.ReadLine();
							if(lastPos == 0)//ヘッダ部をスキップ
							{
	                            lastPos++;
								continue;
							}

				            strResults = line.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

				            if(strResults.Length <= 4)//異常系：成型条件変更等のログの場合、スキップ
				            {
	                            lastPos++;
								continue;
							}
				            
							datetimeValue = strResults[0];
							timeValue = datetimeValue.Substring(datetimeValue.Length - 8);
							
							if(timeValue == lastTime)
							{
								break;
							}
							
							lastPos++;
						}
						
						int indexPos = 0;
						sr.DiscardBufferedData();//一度先頭に戻す
						sr.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
			            while (sr.EndOfStream == false)
			            {
							//データ取得
			                line = sr.ReadLine();

							if (indexPos <= lastPos)//前回の更新部分までスキップ
							{
							    indexPos++;
							    continue;
							}

	                        strResults = line.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

				            if(strResults.Length <= 4)//更新部分にも異常系があれば、スキップ
				            {
								indexPos++;
								continue;
							}

							ParseLogString(strResults);

							indexPos++;
						}
					}
				}
				else if(updateMode == 2)//ログファイルが変わった時→全て読み込む
				{
					int jj = 0;
					sr.DiscardBufferedData();//一度先頭に戻す
					sr.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
		            while (sr.EndOfStream == false)
		            {
		                line = sr.ReadLine();

						if(jj == 0)//ヘッダ部をスキップ
						{
							jj++;
							continue;
						}

			            strResults = line.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

			            if(strResults.Length <= 4)//異常系：成型条件変更等のログの場合、スキップ
			            {
							jj++;
							continue;
						}

						ParseLogString(strResults);

						jj++;
					}
				}

				//一時ファイルを閉じる
	            sr.Close();

				try
				{
					File.Delete(@dstFile);//一時ファイルは削除
				}
				catch (System.IO.IOException ex)
			    {
			        // ファイルを開くのに失敗したときエラーメッセージを表示
					string errorStr = "CSV一時ファイルが削除できなかった可能性があります";
				    System.Console.WriteLine(errorStr);
			        System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
					continue;
			    }
			}
        }

		public void LogFileOut(string logMessage)
		{
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string path = stCurrentDir + "\\PgmMachineOutData.log";
			
			using(var sw = new System.IO.StreamWriter(path, true, System.Text.Encoding.Default))
			{
				sw.WriteLine($"{DateTime.Now.ToLongDateString()} {DateTime.Now.ToLongTimeString()}");
				sw.WriteLine($"  {logMessage}");
				sw.WriteLine ("--------------------------------------------------------------");
			}
		}

		public void ClearListView()
		{
			listView1.Items.Clear();
			//ヘッダの幅を自動調節
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
		}

		public void AddListView()
		{

			if(backSeihinName != currentSeihin)
			{
				backSeihinName = currentSeihin;
				//製品名が変わった時、リストビューをクリア
				ClearListView();
			}

			string[] item1 = {shimeSign, seikei1KataNo, seikei1Tkeisu, seikei1ct1up, seikei1ct1dn, seikei1cc2_cc1, seikei1Zvalue, seikei1Zhosei, 
			/*cc1Value, cc2Value, cc3Value,*/ seikei2cp, seikei1TotalTime, dateValue, timeValue, 
			nikuUpLimit, nikuLoLimit, 
			nikuData[0], nikuResult[0], resultCause[0], 
			nikuData[1], nikuResult[1], resultCause[1], 
			nikuData[2], nikuResult[2], resultCause[2], 
			nikuData[3], nikuResult[3], resultCause[3], 
			nikuData[4], nikuResult[4], resultCause[4], 
			nikuData[5], nikuResult[5], resultCause[5], 
			seikei1TounyuNo, tounyuValue};

			//全体の色設定
            listView1.Sorting = SortOrder.None;
            listView1.ForeColor = Color.Black;//初期の色
            listView1.BackColor = Color.Pink;//全体背景色

			listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加

			DateTime dt3 = DateTime.Parse(timeValue);//タイムスタンプ(文字列)→DateTimeに変換

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

			if(seikei1KataNo == "0" || seikei1KataNo == "")
			{
				listView1.Items[0].BackColor = Color.Gray;
			}

			if(noonSta <= dt3 && dt3 < noonEnd)
			{
				listView1.Items[0].ForeColor = Color.Blue;
			}
			if(sunsetSta <= dt3 && dt3 < sunsetEnd)
			{
				listView1.Items[0].ForeColor = Color.Green;
			}
			if(nightSta <= dt3 && dt3 < nightEnd)
			{
				listView1.Items[0].ForeColor = Color.Purple;
			}

            listView1.Font = new System.Drawing.Font("Times New Roman", 14, System.Drawing.FontStyle.Regular);

			//加圧時間の範囲外なら、強調表示
			//加圧時間の決定
			currentKaatsujikan = currentSeihin;
			for(int i = 0; i < SETDATA.seihinName.Length; i++)
			{
				if(SETDATA.seihinName[i] == currentKaatsujikan)
				{
					if(SETDATA.KaatsuJikanUpper[i] != "" && SETDATA.KaatsuJikanLower[i] != "")
					{
						currentKaatsujikanUp = int.Parse(SETDATA.KaatsuJikanUpper[i]);
						currentKaatsujikanLo = int.Parse(SETDATA.KaatsuJikanLower[i]);
						break;
					}
				}
			}

			numericUpDown3.Text = currentKaatsujikanUp.ToString();
			numericUpDown4.Text = currentKaatsujikanLo.ToString();

			string tmp_cc32Value = seikei1cc2_cc1;
			double kaatsuTime = double.Parse(tmp_cc32Value);
			if(seikei1KataNo != "0" && seikei1KataNo != "")
			{
				if(kaatsuTime < currentKaatsujikanLo || currentKaatsujikanUp < kaatsuTime)
				{
					listView1.Items[0].ForeColor = Color.Yellow;
		            listView1.Items[0].Font = new System.Drawing.Font("Times New Roman", 14, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic);
				}
			}

			//ヘッダの幅を自動調節
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        static private void ReadDataFromXml()
		{
            SETDATA.load(ref Form6.SETDATA);
		}

		static public void WriteDataToXml()
		{
            SETDATA.save(Form6.SETDATA);
		}

		public bool WriteDataToCsv(string logStr, string seihin, string machine, string [] strResults)
		{
            try
            {
		        // appendをtrueにすると，既存のファイルに追記
		        //         falseにすると，ファイルを新規作成する
		        var append = false;
		        // 出力用のファイルを開く
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = currentCsvFile = stCurrentDir + "\\PgmOut_" + seihin + "_" + machine + "_マルチCav.csv";

                string buf = "";
                if(System.IO.File.Exists(path))//既にファイルが存在する
				{
					append = true;
				}
				
		        using(var sw = new System.IO.StreamWriter(path, append, System.Text.Encoding.Default))
		        {
					if(!append)
					{
						buf = string.Format("日付");
						buf += string.Format(",時間");
						buf += string.Format(",成型機");
						buf += string.Format(",製品名");
						buf += string.Format(",ﾛｸﾞﾌｧｲﾙ名");
						buf += string.Format(",締め位置");

						for(int i = 0; i < strResults.Length; i++)
						{
							buf += ',' + string.Format(strResults[i]);
						}

						buf += string.Format(",肉厚上限");
						buf += string.Format(",肉厚下限");
						buf += string.Format(",Cav1測定値");
						buf += string.Format(",Cav1合否");
						buf += string.Format(",Cav1原因");
						buf += string.Format(",Cav2測定値");
						buf += string.Format(",Cav2合否");
						buf += string.Format(",Cav2原因");
						buf += string.Format(",Cav3測定値");
						buf += string.Format(",Cav3合否");
						buf += string.Format(",Cav3原因");
						buf += string.Format(",Cav4測定値");
						buf += string.Format(",Cav4合否");
						buf += string.Format(",Cav4原因");
						buf += string.Format(",Cav5測定値");
						buf += string.Format(",Cav5合否");
						buf += string.Format(",Cav5原因");
						buf += string.Format(",Cav6測定値");
						buf += string.Format(",Cav6合否");
						buf += string.Format(",Cav6原因");
						buf += string.Format(",放射率");
						buf += string.Format(",作業者");
						buf += string.Format(",ｼｮｯﾄ数");
						buf += string.Format(",成型数");

	                    sw.WriteLine(buf);

						buf = "";
		                DateTime dt = DateTime.Now;
						buf = string.Format("{0}", dt.ToString("yyyy/MM/dd"));//日付
						buf += string.Format(",{0}", dt.ToString("HH:mm:ss"));//時間
						buf += logStr;
	                    sw.WriteLine(buf);
					}
					else
					{
		                DateTime dt = DateTime.Now;
						buf = string.Format("{0}", dt.ToString("yyyy/MM/dd"));//日付
						buf += string.Format(",{0}", dt.ToString("HH:mm:ss"));//時間
						buf += logStr;
	                    sw.WriteLine(buf);
					}
		        }
		    }
			catch (System.IO.IOException ex)
		    {
		        // ファイルを開くのに失敗したときエラーメッセージを表示
				string errorStr = "状変時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
				return false;
		    }
		    return true;

		}

		public class SYSSET:System.ICloneable
		{
			public int windowWidth;
			public int windowHeight;
			public int listviewWidth;
			public int listviewHeight;
			public int maxShotCount;

			public string[] goukiName =		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] seihinName=		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] nikuUpper=		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] nikuLower=		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] ngCause=		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] OperatorName=	{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] machineKind = {"", "", "", ""};

            public string[] KaatsuJikanUpper ={"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

            public string[] KaatsuJikanLower ={"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                             "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};


			public int machineType;//5:HSマルチ
			public string selectedMachine = "";
			public string selectedOperator = "";

            public bool load(ref SYSSET ss)
			{
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\PgmDataMonitorSettingData.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamReader fs = new System.IO.StreamReader(path, System.Text.Encoding.Default);
					SYSSET obj;
					obj = (SYSSET)sz.Deserialize(fs);
					fs.Close();
					obj = (SYSSET)obj.Clone();
					ss = obj;
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return(ret);
			}

			public Object Clone()
			{
				SYSSET cln = (SYSSET)this.MemberwiseClone();
				return (cln);
			}

			public bool save(SYSSET ss)
			{
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\PgmDataMonitorSettingData.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamWriter fs = new System.IO.StreamWriter(path, false, System.Text.Encoding.Default);
					sz.Serialize(fs, ss);
					fs.Close();
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return (ret);
			}
		}

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
			listView1.Focus();
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
			listView1.Focus();
        }


        private void button1_Click(object sender, EventArgs e)
        {
			if(!timer1.Enabled)
			{
	            int index = comboBox2.SelectedIndex;//成型機
	            int index2 = comboBox12.SelectedIndex;//作業者
				if(index != -1 && index2 != -1)
				{
					timer1.Enabled = true;//監視タイマー開始
					button1.Enabled = false;
				}
				else
				{
		            MessageBox.Show("成型機と作業者を選択すれば開始します", "PGM成型機監視アプリ", MessageBoxButtons.OK);
				}
			}
        }

        private void Form6_FormClosed(object sender, FormClosedEventArgs e)
        {
			string errorStr = "正常に終了しました。";
			LogFileOut(errorStr);

            Application.Exit();
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
		    if(e.Button == System.Windows.Forms.MouseButtons.Right)
			{
				return;
			}

			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
				timer1.Enabled = false;

			    int index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置

				if(listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "")
				{
					timer1.Enabled = true;
					return;
				}

				int j = 0;
				nikuEdit = causeEdit = "";
				for(int i = 0; i < nikuData.Length; i++)
				{
					nikuData[i] = listView1.Items[index].SubItems[14 + j].Text;//肉厚入力値
					nikuResult[i] = listView1.Items[index].SubItems[15 + j].Text;//合否
					resultCause[i] = listView1.Items[index].SubItems[16 + j].Text;//不良原因

					if(i == 0)
					{
						nikuEdit += nikuData[i];
						causeEdit += resultCause[i];
					}
					else
					{
						nikuEdit += "," + nikuData[i];
						causeEdit += "," + resultCause[i];
					}
					
					j += 3;
				}

                string justSleeve = listView1.Items[index].SubItems[1].Text;
				nikuUpLimit = listView1.Items[index].SubItems[12].Text;//肉厚上限
				nikuLoLimit = listView1.Items[index].SubItems[13].Text;//肉厚下限
				string currentShotCount = listView1.Items[index].SubItems[32].Text;//ショット数

                //登録・編集画面の表示
                Form7 form7 = new Form7();
				form7.SetInfo(justSleeve, currentShotCount, nikuUpLimit, nikuLoLimit, nikuEdit, causeEdit);
	            form7.ShowDialog();

	            string InfoStr = form7.EditInfo;
	          	string CauseStr = form7.CauseInfo;
				string SeikeiStr = form7.SeikeiInfo;
	            if(InfoStr == "")
	            {
					timer1.Enabled = true;
					return;
				}

				string[] strline = InfoStr.Split(',');

				double nikuUpData = double.Parse(nikuUpLimit);
				double nikuLoData = double.Parse(nikuLoLimit);

				string[] causeline = CauseStr.Split(',');
				string[] seikeiline = SeikeiStr.Split(',');

				//CSV更新:別ファイルに全て読んで、一部を書き換えてファイル名を変える
                string line = "";
                StreamReader reader = null;
                StreamWriter writer = null;
                string path = "";
				try
				{
	                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
                    string[] lines = File.ReadAllLines(currentCsvFile);
                    int lineMax = lines.Length;//CSVの行数取得

                    path = currentCsvFile + ".tmp";
					writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

					int count = 0;
					index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置
					while(reader.Peek() >= 0)
					{
						line = reader.ReadLine();

						if((lineMax - 1) - index == count)
						{
							string[] cols = line.Split(',');
							string buf = "";
							for(int i = 0; i < cols.Length; i++)
							{
								if(i == 0)
								{
									buf = cols[i];
								}
								else if(i == 57)//肉厚上限
								{
									buf += "," + nikuUpLimit;
								}
								else if(i == 58)//肉厚下限
								{
									buf += "," + nikuLoLimit;
								}

								else if(i == 59)//Cav1肉厚値
								{
									buf += "," + strline[0];
								}
								else if(i == 60)//Cav1合否
								{
									if(strline[0] == "-")
									{
										buf += ",-";
									}
									else
									{
										if(causeline[0] == "なし")
										{
											buf += "," + "OK";
										}
										else
										{
											buf += "," + "NG";
										}
									}
								}
								else if(i == 61)//Cav1不良原因
								{
									if(causeline[0] == "なし")
									{
										buf += ",-";
									}
									else
									{
										buf += "," + causeline[0];
									}
								}

								else if(i == 62)//Cav2肉厚値
								{
									buf += "," + strline[1];
								}
								else if(i == 63)//Cav2合否
								{
									if(strline[1] == "-")
									{
										buf += ",-";
									}
									else
									{
										if(causeline[1] == "なし")
										{
											buf += "," + "OK";
										}
										else
										{
											buf += "," + "NG";
										}
									}
								}
								else if(i == 64)//Cav2不良原因
								{
									if(causeline[1] == "なし")
									{
										buf += ",-";
									}
									else
									{
										buf += "," + causeline[1];
									}
								}

								else if(i == 65)//Cav3肉厚値
								{
									buf += "," + strline[2];
								}
								else if(i == 66)//Cav3合否
								{
									if(strline[2] == "-")
									{
										buf += ",-";
									}
									else
									{
										if(causeline[2] == "なし")
										{
											buf += "," + "OK";
										}
										else
										{
											buf += "," + "NG";
										}
									}
								}
								else if(i == 67)//Cav3不良原因
								{
									if(causeline[2] == "なし")
									{
										buf += ",-";
									}
									else
									{
										buf += "," + causeline[2];
									}
								}

								else if(i == 68)//Cav4肉厚値
								{
									buf += "," + strline[3];
								}
								else if(i == 69)//Cav4合否
								{
									if(strline[3] == "-")
									{
										buf += ",-";
									}
									else
									{
										if(causeline[3] == "なし")
										{
											buf += "," + "OK";
										}
										else
										{
											buf += "," + "NG";
										}
									}
								}
								else if(i == 70)//Cav4不良原因
								{
									if(causeline[3] == "なし")
									{
										buf += ",-";
									}
									else
									{
										buf += "," + causeline[3];
									}
								}

								else if(i == 71)//Cav5肉厚値
								{
									buf += "," + strline[4];
								}
								else if(i == 72)//Cav5合否
								{
									if(strline[4] == "-")
									{
										buf += ",-";
									}
									else
									{
										if(causeline[4] == "なし")
										{
											buf += "," + "OK";
										}
										else
										{
											buf += "," + "NG";
										}
									}
								}
								else if(i == 73)//Cav5不良原因
								{
									if(causeline[4] == "なし")
									{
										buf += ",-";
									}
									else
									{
										buf += "," + causeline[4];
									}
								}

								else if(i == 74)//Cav6肉厚値
								{
									buf += "," + strline[5];
								}
								else if(i == 75)//Cav6合否
								{
									if(strline[5] == "-")
									{
										buf += ",-";
									}
									else
									{
										if(causeline[5] == "なし")
										{
											buf += "," + "OK";
										}
										else
										{
											buf += "," + "NG";
										}
									}
								}
								else if(i == 76)//Cav6不良原因
								{
									if(causeline[5] == "なし")
									{
										buf += ",-";
									}
									else
									{
										buf += "," + causeline[5];
									}
								}
								else if(i == 79)//ショット数
								{
									buf += "," + seikeiline[0];
								}
								else
								{
									buf += "," + cols[i];
								}
							}
							writer.WriteLine(buf);
						}
						else
						{
							writer.WriteLine(line);
						}

						count++;
					}

	                //画面:ListViewの更新
					listView1.Items[index].SubItems[12].Text = nikuUpLimit;//肉厚上限
					listView1.Items[index].SubItems[13].Text = nikuLoLimit;//肉厚下限
					listView1.Items[index].SubItems[32].Text = seikeiline[0];//ショット数

					listView1.Items[index].SubItems[14].Text = string.Format("{0:#.###}", strline[0]);
					listView1.Items[index].SubItems[17].Text = string.Format("{0:#.###}", strline[1]);
					listView1.Items[index].SubItems[20].Text = string.Format("{0:#.###}", strline[2]);
					listView1.Items[index].SubItems[23].Text = string.Format("{0:#.###}", strline[3]);
					listView1.Items[index].SubItems[26].Text = string.Format("{0:#.###}", strline[4]);
					listView1.Items[index].SubItems[29].Text = string.Format("{0:#.###}", strline[5]);

					listView1.Items[0].UseItemStyleForSubItems = false;
 					int jj = 0;
					for(int i = 0; i < causeline.Length; i++)
					{
						if(strline[i] == "-")
						{
							listView1.Items[index].SubItems[15 + jj].Text = "-";
							listView1.Items[index].SubItems[16 + jj].Text = "-";

							listView1.Items[index].SubItems[14 + jj].BackColor = Color.White;
							listView1.Items[index].SubItems[15 + jj].BackColor = Color.White;
							listView1.Items[index].SubItems[16 + jj].BackColor = Color.White;
						}
						else
						{
							if(causeline[i] == "なし")
							{
								listView1.Items[index].SubItems[15 + jj].Text = "OK";
								listView1.Items[index].SubItems[16 + jj].Text = "-";

								listView1.Items[index].SubItems[14 + jj].BackColor = Color.Lime;
								listView1.Items[index].SubItems[15 + jj].BackColor = Color.Lime;
								listView1.Items[index].SubItems[16 + jj].BackColor = Color.Lime;
							}
							else
							{
								listView1.Items[index].SubItems[15 + jj].Text = "NG";
								listView1.Items[index].SubItems[16 + jj].Text = causeline[i];

								listView1.Items[index].SubItems[14 + jj].BackColor = Color.Red;
								listView1.Items[index].SubItems[15 + jj].BackColor = Color.Red;
								listView1.Items[index].SubItems[16 + jj].BackColor = Color.Red;
							}
						}
						jj += 3;
					}

					//ヘッダの幅を自動調節
					listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
				}
				catch (System.IO.IOException ex)
				{
					string errorStr = "肉厚入力時にCSVファイルを開けなかった可能性があります";
				    System.Console.WriteLine(errorStr);
			        System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
				}
				finally
				{
					if(reader != null)
					{
						reader.Close();
						//元ファイル削除
						File.Delete(@currentCsvFile);
					}
					if(writer != null)
					{
						writer.Close();
						//一時ファイル→元ファイルへファイル名変更
						System.IO.File.Move(@path, @currentCsvFile);
					}

				}

				timer1.Enabled = true;
			}

        }

        private void Form6_Shown(object sender, EventArgs e)
        {
//			if(formCalender == null || formCalender.IsDisposed)
//			{
//				formCalender = new FormCalender();
//                formCalender.Show();
//				formCalender.SetParentForm(this, SETDATA.machineType);
//            }
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            if (hidden == 0)
            {
                Visible = true;
                hidden = 1;
            }
            else
            {
                Visible = false;
                hidden = 0;
            }
        }

        private void Form6_FormClosing(object sender, FormClosingEventArgs e)
        {
			string mes = "アプリを終了すると成型機からのデータを受け取れなくなります！" + "\r\n" + "本当に終了しますか？";
			DialogResult result = MessageBox.Show(mes, "PGM成型機監視アプリ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
			if(result == DialogResult.Yes)
			{
//	            MessageBox.Show("作業が終わったら、すぐ起動して下さい！", "PGM成型機監視アプリ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			else
			{
                e.Cancel = true;
			}
        }

        private void button2_Click(object sender, EventArgs e)
        {
			if(formCalender == null || formCalender.IsDisposed)
			{
				formCalender = new FormCalender();
                formCalender.Show();
				formCalender.SetParentForm(this, SETDATA.machineType);
            }
        }

		public void GetDisplayData(ref string gouki, ref string seihin, ref string upper, ref string lower, ref string operatorName)
		{
			gouki = comboBox2.Text;
			seihin = label1.Text;
			upper = numericUpDown12.Text;
			lower = numericUpDown13.Text;
            operatorName = comboBox12.Text;
		}

        public void SetDateTimeToAnalize(string start, string end, ref string okngStr, ref string ngcauseStr, ref string nikuInfo, ref string sleeveInfo, ref double kadou, 
										ref int total, 
										ref string ref_nikuatsuNG, ref string ref_kizuNG, ref string ref_butsuNG, ref string ref_yakeNG, 
										ref string ref_hibicrackNG, ref string ref_gasukizuNG, ref string ref_houshakizuNG, 
										ref string ref_giratsukikumoriNG, ref string ref_hennikumendareNG, ref string ref_yogoreNG, 
										ref string ref_hokoriNG, ref string ref_keijoseidoNG, ref string ref_etcNG
		)
		{
            DateTime dtsta = DateTime.Parse(start);
            DateTime dtend = DateTime.Parse(end);

	        int okCount = 0;
	        int ngCount = 0;

			int allSleeve = 0;
			int workSleeve = 0;
			int nikuatsuNG = 0;//肉厚不良
			int kizuNG = 0;//キズ
			int butsuNG = 0;//ブツ
			int yakeNG = 0;//ヤケ
			int hibicrackNG = 0;//ヒビ/クラック
			int gasukizuNG = 0;//ガスキズ
			int houshakizuNG = 0;//放射キズ
			int giratsukikumoriNG = 0;//ギラツキ/クモリ
			int hennikumendareNG = 0;//偏肉/面ダレ
			int yogoreNG = 0;//汚れ
			int hokoriNG = 0;//ほこり
			int keijoseidoNG = 0;//形状精度
			int etcNG = 0;//その他

            double nikuData = 0;
            double maxNiku = 0;
            double minNiku = 100;
            double sumNiku = 0;
            double sumV_Niku = 0;
			int nikuCount = 0;
			var nikuList = new List<double>();
			List<SleeveList> list = new List<SleeveList>();

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                int colmax = listView1.Items[i].SubItems.Count;
                string listDate = "";
                string listTime = "";
                string listDT = "";
                string [] result = new string [6];
                string [] ngcause = new string [6];
                string shotCount = "";
                string seikeiCount = "";
				string [] nikuatsuSokutei = new string [6];
				string sleeveNumber = "";
                for (int j = 0; j < colmax; j++)
                {           
                    string listitem = listView1.Items[i].SubItems[j].Text;
                    if (j == 10)//日付
                    {
                        listDate = listitem;
                    }
                    else if(j == 11)//時間
                    {
                        listTime = listitem;
                    }
                    else if(j == 14 || j == 17 || j == 20 || j == 23 || j == 26 || j == 29)//肉厚測定値
					{
						nikuatsuSokutei[(j - 14) / 3] = listitem;
					}
                    else if(j == 15 || j == 18 || j == 21 || j == 24 || j == 27 || j == 30)//合否
					{
						result[(j - 14) / 3] = listitem;
					}
                    else if(j == 16 || j == 19 || j == 22 || j == 25 || j == 28 || j == 31)//不良原因
					{
						ngcause[(j - 14) / 3] = listitem;
					}
                    else if(j == 1)//スリーブ番号
					{
						sleeveNumber = listitem;
					}
                    else if(j == 32)//ショット数
					{
						shotCount = listitem;
					}
                    else if(j == 33)//成型数
					{
						seikeiCount = listitem;
					}

                }

                listDT = listDate + " " + listTime;
                DateTime dttarget = DateTime.Parse(listDT);
                if(dtsta <= dttarget && dttarget <= dtend)
                {
					allSleeve++;

	                for (int j = 0; j < result.Length; j++)
					{
						if(result[j] == "OK")
						{
		                    okCount++;
		                }
						else if(result[j] == "NG")
						{
		                    ngCount++;

							if(ngcause[j] == "肉厚不良")
							{
								nikuatsuNG++;//肉厚不良
							}
							else if(ngcause[j] == "キズ")
							{
								kizuNG++;//キズ
							}
							else if(ngcause[j] == "ブツ")
							{
								butsuNG++;//ブツ
							}
							else if(ngcause[j] == "ヤケ")
							{
								yakeNG++;//ヤケ
							}
							else if(ngcause[j] == "ヒビ/クラック")
							{
								hibicrackNG++;//ヒビ/クラック
							}
							else if(ngcause[j] == "ガスキズ")
							{
								gasukizuNG++;//ガスキズ
							}
							else if(ngcause[j] == "放射キズ")
							{
								houshakizuNG++;//放射キズ
							}
							else if(ngcause[j] == "ギラツキ/クモリ")
							{
								giratsukikumoriNG++;//ギラツキ/クモリ
							}
							else if(ngcause[j] == "偏肉/面ダレ")
							{
								hennikumendareNG++;//偏肉/面ダレ
							}
							else if(ngcause[j] == "汚れ")
							{
								yogoreNG++;//汚れ
							}
							else if(ngcause[j] == "ほこり")
							{
								hokoriNG++;//ほこり
							}
							else if(ngcause[j] == "形状精度")
							{
								keijoseidoNG++;//形状精度
							}
							else if(ngcause[j] == "その他")
							{
								etcNG++;//その他
							}
						}
	                
		                if(nikuatsuSokutei[j] != "-" && nikuatsuSokutei[j] != "")
						{
			                nikuData = double.Parse(nikuatsuSokutei[j].ToString());
			                sumNiku += nikuData;
			                sumV_Niku += (nikuData * nikuData);
			                nikuList.Add(nikuData);
			             	nikuCount++;
			                
			                if(nikuData > maxNiku)
			                {
								maxNiku = nikuData;
							}
			                if(nikuData < minNiku)
			                {
								minNiku = nikuData;
							}
						}
	                }

					if(sleeveNumber != "0" && sleeveNumber != "")
					{
						workSleeve++;

						//未登録なら追加
	                    if(list.Find(m => m.sleeveNumber == sleeveNumber).sleeveNumber != sleeveNumber)
		                {
	                        SleeveList sl = new SleeveList();
	                        sl.sleeveNumber= sleeveNumber;
							sl.shotCount = int.Parse(shotCount);

			                for (int j = 0; j < result.Length; j++)
							{
								if(result[j] == "OK")
								{
									sl.ryohinSuu++;
								}
								else if(result[j] == "NG")
								{
									sl.furyouSuu++;
								}
							}
							//SL毎の不良内訳
			                for (int j = 0; j < ngcause.Length; j++)
							{
								if(ngcause[j] == "肉厚不良")
								{
									sl.nikuatsuNG++;
								}
								else if(ngcause[j] == "キズ")
								{
									sl.kizuNG++;
								}
								else if(ngcause[j] == "ブツ")
								{
									sl.butsuNG++;
								}
								else if(ngcause[j] == "ヤケ")
								{
									sl.yakeNG++;
								}
								else if(ngcause[j] == "ヒビ/クラック")
								{
									sl.hibicrackNG++;
								}
								else if(ngcause[j] == "ガスキズ")
								{
									sl.gasukizuNG++;
								}
								else if(ngcause[j] == "放射キズ")
								{
									sl.houshakizuNG++;
								}
								else if(ngcause[j] == "ギラツキ/クモリ")
								{
									sl.giratsukikumoriNG++;
								}
								else if(ngcause[j] == "偏肉/面ダレ")
								{
									sl.hennikumendareNG++;
								}
								else if(ngcause[j] == "汚れ")
								{
									sl.yogoreNG++;
								}
								else if(ngcause[j] == "ほこり")
								{
									sl.hokoriNG++;
								}
								else if(ngcause[j] == "形状精度")
								{
									sl.keijoseidoNG++;
								}
								else if(ngcause[j] == "その他")
								{
									sl.etcNG++;
								}
							}

							list.Add(sl);
						}
		                else
		                {
							for(int j = 0; j < list.Count; j++)
							{
								if(list[j].sleeveNumber == sleeveNumber)
								{
									int sc = int.Parse(shotCount);
									if(list[j].shotCount < sc)
									{
										SleeveList tmpList = list[j];
										tmpList.shotCount = sc;
										list[j] = tmpList;
									}

									SleeveList causeList = list[j];

					                for (int k = 0; k < result.Length; k++)
									{
										if(result[k] == "OK")
										{
											causeList.ryohinSuu++;
										}
										else if(result[k] == "NG")
										{
											causeList.furyouSuu++;
										}
									}

									//SL毎の不良内訳
					                for (int k = 0; k < result.Length; k++)
									{
										if(ngcause[k] == "肉厚不良")
										{
											causeList.nikuatsuNG++;
										}
										else if(ngcause[k] == "キズ")
										{
											causeList.kizuNG++;
										}
										else if(ngcause[k] == "ブツ")
										{
											causeList.butsuNG++;
										}
										else if(ngcause[k] == "ヤケ")
										{
											causeList.yakeNG++;
										}
										else if(ngcause[k] == "ヒビ/クラック")
										{
											causeList.hibicrackNG++;
										}
										else if(ngcause[k] == "ガスキズ")
										{
											causeList.gasukizuNG++;
										}
										else if(ngcause[k] == "放射キズ")
										{
											causeList.houshakizuNG++;
										}
										else if(ngcause[k] == "ギラツキ/クモリ")
										{
											causeList.giratsukikumoriNG++;
										}
										else if(ngcause[k] == "偏肉/面ダレ")
										{
											causeList.hennikumendareNG++;
										}
										else if(ngcause[k] == "汚れ")
										{
											causeList.yogoreNG++;
										}
										else if(ngcause[k] == "ほこり")
										{
											causeList.hokoriNG++;
										}
										else if(ngcause[k] == "形状精度")
										{
											causeList.keijoseidoNG++;
										}
										else if(ngcause[k] == "その他")
										{
											causeList.etcNG++;
										}
									}
									
									list[j] = causeList;

								}
							}
						}
					}
	                
                }

            }


			//一括入力のComboBoxのSLを再登録
			comboBox1.Items.Clear();
			for(int i = 0; i < list.Count; i++)
			{
				comboBox1.Items.Add(list[i].sleeveNumber);
			}

            //OK,NGの数を結合する
            okngStr = okCount.ToString() + "," + ngCount.ToString();
            //各不良原因の数を結合する
			ngcauseStr = nikuatsuNG + "," + 
						 kizuNG + "," + 
						 butsuNG + "," + 
						 yakeNG + "," + 
						 hibicrackNG + "," + 
						 gasukizuNG + "," + 
						 houshakizuNG + "," + 
						 giratsukikumoriNG + "," + 
						 hennikumendareNG + "," + 
						 yogoreNG + "," + 
						 hokoriNG + "," + 
						 keijoseidoNG + "," + 
						 etcNG;

			//標準偏差を求める
			double mean = sumNiku / nikuCount;
			double variance = (sumV_Niku / nikuCount) - (mean * mean);
			double stddev = Math.Sqrt(variance);

			//最大肉厚、最小肉厚、平均肉厚、標準偏差を結合する
			if(nikuCount > 0)
			{
				double aveNiku = (double)(sumNiku / nikuCount);
				nikuInfo = maxNiku.ToString("F3") + "," + minNiku.ToString("F3") + "," + aveNiku.ToString("F3") + "," + stddev.ToString("F3");
			}

			//スリーブ番号、ショット数を結合する
			int jj = 0;
			for(int i = 0; i < list.Count; i++)
			{
				int seikeiTotal = list[i].ryohinSuu + list[i].furyouSuu;
				if(jj == 0)
				{
					sleeveInfo += list[i].sleeveNumber + "," + list[i].shotCount.ToString() + "," + seikeiTotal.ToString() + "," + list[i].ryohinSuu.ToString() + "," + list[i].furyouSuu.ToString();
				}
				else
				{
					sleeveInfo += "," + list[i].sleeveNumber + "," + list[i].shotCount.ToString() + "," + seikeiTotal.ToString() + "," + list[i].ryohinSuu.ToString() + "," + list[i].furyouSuu.ToString();
				}
				jj++;
			}
			
			//スリーブの稼働率
			kadou = (double)((double)workSleeve / (double)allSleeve);


			total = list.Count;
			for(int j = 0; j < list.Count; j++)
			{
				if(j == 0)
				{
					ref_nikuatsuNG = list[j].nikuatsuNG.ToString();
					ref_kizuNG = list[j].kizuNG.ToString();
					ref_butsuNG = list[j].butsuNG.ToString();
					ref_yakeNG = list[j].yakeNG.ToString();
					ref_hibicrackNG = list[j].hibicrackNG.ToString();
					ref_gasukizuNG = list[j].gasukizuNG.ToString();
					ref_houshakizuNG = list[j].houshakizuNG.ToString();
					ref_giratsukikumoriNG = list[j].giratsukikumoriNG.ToString();
					ref_hennikumendareNG = list[j].hennikumendareNG.ToString();
					ref_yogoreNG = list[j].yogoreNG.ToString();
					ref_hokoriNG = list[j].hokoriNG.ToString();
					ref_keijoseidoNG = list[j].keijoseidoNG.ToString();
					ref_etcNG = list[j].etcNG.ToString();
					continue;
				}

				ref_nikuatsuNG += "," + list[j].nikuatsuNG.ToString();
				ref_kizuNG += "," + list[j].kizuNG.ToString();
				ref_butsuNG += "," + list[j].butsuNG.ToString();
				ref_yakeNG += "," + list[j].yakeNG.ToString();
				ref_hibicrackNG += "," + list[j].hibicrackNG.ToString();
				ref_gasukizuNG += "," + list[j].gasukizuNG.ToString();
				ref_houshakizuNG += "," + list[j].houshakizuNG.ToString();
				ref_giratsukikumoriNG += "," + list[j].giratsukikumoriNG.ToString();
				ref_hennikumendareNG += "," + list[j].hennikumendareNG.ToString();
				ref_yogoreNG += "," + list[j].yogoreNG.ToString();
				ref_hokoriNG += "," + list[j].hokoriNG.ToString();
				ref_keijoseidoNG += "," + list[j].keijoseidoNG.ToString();
				ref_etcNG += "," + list[j].etcNG.ToString();
			}

        }

		public void GetShimeHani(DateTime startTime, DateTime endTime, ref string staDate, ref string endDate)
		{
			string start_sign = startTime.ToString("MMdd");
			string end_sign = endTime.ToString("MMdd");
			string sta = start_sign + "開始";
			string end = end_sign + "終了";

            for (int i = 0; i < listView1.Items.Count; i++)
            {
				string listitem = listView1.Items[i].SubItems[0].Text;
				if(listitem == sta)
				{
					staDate = listView1.Items[i].SubItems[10].Text + " " + listView1.Items[i].SubItems[11].Text;
				}
				else if(listitem == end)
				{
					endDate = listView1.Items[i].SubItems[10].Text + " " + listView1.Items[i].SubItems[11].Text;
				}
			}
		}

        private void listView1_MouseUp(object sender, MouseEventArgs e)
        {
		    if(e.Button == System.Windows.Forms.MouseButtons.Right)
		    {
				if(listView1.SelectedItems.Count > 0)
				{
				    int index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置

					System.Drawing.Point p = System.Windows.Forms.Cursor.Position;
					this.contextMenuStrip1.Show(p);
				}
		    }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if(listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "")
			{
				return;
			}

			timer1.Enabled = false;

			DateTime dt = DateTime.Now;
			string sign = dt.ToString("MMdd");
			string sta = sign + "開始";//締め開始指定

			if(listView1.Items[index].SubItems[0].Text == "")
			{
				SetShimeSign(sta);
			}

			timer1.Enabled = true;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if(listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "")
			{
				return;
			}

			timer1.Enabled = false;

			DateTime dt = DateTime.Now;
			string sign = dt.ToString("MMdd");
			string end = sign + "終了";//締め終了指定

			if(listView1.Items[index].SubItems[0].Text == "")
			{
				SetShimeSign(end);
			}

			timer1.Enabled = true;
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if(listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "")
			{
				return;
			}

			timer1.Enabled = false;

			DateTime dt = DateTime.Now;
			dt = dt.AddDays(1);
			string sign = dt.ToString("MMdd");
			string sta = sign + "開始";//締め開始指定

			if(listView1.Items[index].SubItems[0].Text == "")
			{
				SetShimeSign(sta);
			}

			timer1.Enabled = true;
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if(listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "")
            {
                return;
            }

            timer1.Enabled = false;

			if(listView1.Items[index].SubItems[0].Text != "")
			{
				SetShimeSign("");
			}

            timer1.Enabled = true;
        }

        private void ResetAllShimeSign(string sta, string end, string next_sta, string next_end)
        {
			//CSV更新:別ファイルに全て読んで、一部を書き換えてファイル名を変える
            string line = "";
            StreamReader reader = null;
            StreamWriter writer = null;
            string path = "";
			try
			{
                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
                string[] lines = File.ReadAllLines(currentCsvFile);
                path = currentCsvFile + ".tmp";
				writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

				while(reader.Peek() >= 0)
				{
					line = reader.ReadLine();

					string[] cols = line.Split(',');
					if(cols[5] == sta || cols[5] == end || cols[5] == next_sta || cols[5] == next_end)
					{
						string buf = "";
						for(int i = 0; i < cols.Length; i++)
						{
							if(i == 0)
							{
								buf = cols[i];
							}
							else if(i == 5)
							{
								buf += ",";
							}
							else
							{
								buf += "," + cols[i];
							}
						}
						writer.WriteLine(buf);
					}
					else
					{
						writer.WriteLine(line);
					}
					
				}
				
				for(int i = 0; i < listView1.Items.Count; i++)
				{
					if(listView1.Items[i].SubItems[0].Text == sta || listView1.Items[i].SubItems[0].Text == end ||
					listView1.Items[i].SubItems[0].Text == next_sta || listView1.Items[i].SubItems[0].Text == next_end)
					{
						listView1.Items[i].SubItems[0].Text = "";
					}
				}
				
                //ヘッダの幅を自動調節
                listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
			}
			catch (System.IO.IOException ex)
			{
				string errorStr = "締めの更新時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			finally
			{
				if(reader != null)
				{
					reader.Close();
					//元ファイル削除
					File.Delete(@currentCsvFile);
				}
				if(writer != null)
				{
					writer.Close();
					//一時ファイル→元ファイルへファイル名変更
					System.IO.File.Move(@path, @currentCsvFile);
				}

			}
		}

        private void SetShimeSign(string sign)
        {
			//CSV更新:別ファイルに全て読んで、一部を書き換えてファイル名を変える
            string line = "";
            StreamReader reader = null;
            StreamWriter writer = null;
            string path = "";
			try
			{
                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
                string[] lines = File.ReadAllLines(currentCsvFile);
                int lineMax = lines.Length;//CSVの行数取得

                path = currentCsvFile + ".tmp";
				writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

				int count = 0;
				int index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置
				while(reader.Peek() >= 0)
				{
					line = reader.ReadLine();

					if((lineMax - 1) - index == count)
					{
						string[] cols = line.Split(',');
						string buf = "";
						for(int i = 0; i < cols.Length; i++)
						{
							if(i == 0)
							{
								buf = cols[i];
							}
							else if(i == 5)
							{
								buf += "," + sign;
							}
							else
							{
								buf += "," + cols[i];
							}
						}
						writer.WriteLine(buf);
					}
					else
					{
						writer.WriteLine(line);
					}
					
					count++;
				}
                //ヘッダの幅を自動調節
				listView1.Items[index].SubItems[0].Text = sign;
                listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
			}
			catch (System.IO.IOException ex)
			{
				string errorStr = "肉厚入力時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			finally
			{
				if(reader != null)
				{
					reader.Close();
					//元ファイル削除
					File.Delete(@currentCsvFile);
				}
				if(writer != null)
				{
					writer.Close();
					//一時ファイル→元ファイルへファイル名変更
					System.IO.File.Move(@path, @currentCsvFile);
				}

			}
		}

		public bool WriteDailySummaryToCsv(DateTime sta_datehani, DateTime end_datehani, string logStr)
		{
			string sta_sign = sta_datehani.ToString("MMdd");
			string end_sign = end_datehani.ToString("MMdd");
			string sta = sta_sign + "開始";
			string end = end_sign + "終了";
			int staPos = -1;
			int endPos = -1;

			//締めのサインがあるかチェック
            for (int i = 0; i < listView1.Items.Count; i++)
            {
				string listitem = listView1.Items[i].SubItems[0].Text;
				if(listitem == sta)
				{
					staPos = i;
				}
				if(listitem == end)
				{
					endPos = i;
				}

			}

			if(staPos == -1 || endPos == -1)
			{
				return false;
			}

			DateTime dt = DateTime.Now;
			string dailysign = dt.ToString("MMddHHmmss");

            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string dailyDir = stCurrentDir + "\\daily";
            string dailyPath = dailyDir + "\\" + currentSeihin + "_" + currentMachine + "_マルチ_" + sta_sign + "to" + end_sign + "_" + dailysign + ".csv";

			StreamWriter writer = null;

			//dailyフォルダが存在していなければ作成
			if(!Directory.Exists(dailyDir))
			{
				Directory.CreateDirectory(dailyDir);
			}

            try
            {
		        writer = new StreamWriter(dailyPath, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

				string buf = "";
				buf = string.Format("日付");
				buf += string.Format(",号機");
				buf += string.Format(",製品名");
				buf += string.Format(",歩留り");
				buf += string.Format(",成型数");
				buf += string.Format(",良品数");
				buf += string.Format(",不良数");
				buf += string.Format(",肉厚不足");
				buf += string.Format(",キズ");
				buf += string.Format(",ブツ");
				buf += string.Format(",ヤケ");
				buf += string.Format(",ﾋﾋﾞ/ｸﾗｯｸ");
				buf += string.Format(",ｶﾞｽｷｽﾞ");
				buf += string.Format(",放射ｷｽﾞ");
				buf += string.Format(",ｷﾞﾗﾂｷ/ｸﾓﾘ");
				buf += string.Format(",偏肉/面ﾀﾞﾚ");
				buf += string.Format(",汚れ");
				buf += string.Format(",ほこり");
				buf += string.Format(",形状精度");
				buf += string.Format(",その他");
				buf += string.Format(",作業者");
				buf += string.Format(",肉厚平均値");
				buf += string.Format(",肉厚最大値");
				buf += string.Format(",肉厚最小値");
				buf += string.Format(",肉厚標準偏差");

                writer.WriteLine(buf);

				buf = "";
				buf += logStr;
                writer.WriteLine(buf);

		    }
			catch (System.IO.IOException ex)
			{
				string errorStr = "dailyファイルを保存できなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			finally
			{
				if(writer != null)
				{
					writer.Close();
				}

			}
			return true;

		}

		public bool WriteDailyDataToCsv(DateTime sta_datehani, DateTime end_datehani)
		{
			DateTime dt = DateTime.Now;
			string dailysign = dt.ToString("MMddHHmmss");
			string sta_sign = sta_datehani.ToString("MMdd");
			string end_sign = end_datehani.ToString("MMdd");
			string sta = sta_sign + "開始";
			string end = end_sign + "終了";

			//CSV更新:別ファイルに全て読んで、一部を書き換えてファイル名を変える
	        string line = "";
	        StreamReader reader = null;
	        StreamWriter writer = null;
			try
			{
	            reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
	            string[] lines = File.ReadAllLines(currentCsvFile);
	            int lineMax = lines.Length;//CSVの行数取得


				int count = 0;
				int staPos = 0;
				int endPos = 0;
				while(reader.Peek() >= 0)//書き込む場所を探す
				{
					line = reader.ReadLine();

					if(count == 0)
					{
						count++;
						continue;
					}

					string[] cols = line.Split(',');
					if(cols[5] == sta)
					{
						staPos = count;
					}
					else if(cols[5] == end)
					{
						endPos = count;
					}
					count++;
				}

				if(staPos == 0 || endPos == 0)//締めのマークが無い場合
				{
					if(reader != null)
					{
						reader.Close();
					}
					return false;
				}

				reader.BaseStream.Seek(0, SeekOrigin.Begin);//先頭に戻す
				reader.DiscardBufferedData();

                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
				string dailyDir = stCurrentDir + "\\daily";
                string dailyPath = dailyDir + "\\" + currentSeihin + "_" + currentMachine + "_" + "マルチ" + "_" + sta_sign + "to" + end_sign + "_" + dailysign + ".csv";

				//dailyフォルダが存在していなければ作成
				if(!Directory.Exists(dailyDir))
				{
					Directory.CreateDirectory(dailyDir);
				}

				writer = new StreamWriter(dailyPath, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

				count = 0;
				while(reader.Peek() >= 0)
				{
					line = reader.ReadLine();

					if(count == 0)
					{
						writer.WriteLine(line);
						count++;
						continue;
					}

					if(count < staPos)
					{
						count++;
						continue;
					}
					else if(count > endPos)
					{
						count++;
						continue;
					}

					writer.WriteLine(line);
					count++;
				}

			}
			catch (System.IO.IOException ex)
			{
				string errorStr = "肉厚入力時にdayファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			finally
			{
				if(reader != null)
				{
					reader.Close();
				}
				if(writer != null)
				{
					writer.Close();
				}

			}
			return true;
		}

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox1.Checked)
			{
				comboBox1.Enabled = true;
				comboBox3.Enabled = true;
				numericUpDown1.Enabled = true;
                button3.Enabled = true;
            }
            else
			{
				comboBox1.Enabled = false;
				comboBox3.Enabled = false;
				numericUpDown1.Enabled = false;
				button3.Enabled = false;
				comboBox1.Text = "";
				comboBox3.Text = "";
				numericUpDown1.Text = "1";
			}

			List<SleeveList> list = new List<SleeveList>();
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                int colmax = listView1.Items[i].SubItems.Count;
                for (int j = 0; j < colmax; j++)
                {           
                    string listitem = listView1.Items[i].SubItems[j].Text;
                    if(j == 1)//スリーブ番号
					{
						string sleeveNumber = listitem;
						if(sleeveNumber != "0" && sleeveNumber != "")
						{
		                    if(list.Find(m => m.sleeveNumber == sleeveNumber).sleeveNumber != sleeveNumber)
			                {
		                        SleeveList sl = new SleeveList();
		                        sl.sleeveNumber = sleeveNumber;
								list.Add(sl);
							}
						}
					}
				}
			}

			//一括入力のComboBoxのSLを再登録
			comboBox1.Items.Clear();
			for(int i = 0; i < list.Count; i++)
			{
				comboBox1.Items.Add(list[i].sleeveNumber);
			}
        }

        private void button3_Click(object sender, EventArgs e)
        {
			if(comboBox1.Text == "" || comboBox3.Text == "")
			{
				MessageBox.Show("スリーブと不具合原因を選んで下さい", "確認", MessageBoxButtons.OK);
				return;
			}

			string selectSleeve = comboBox1.Text;
			string selectCause = comboBox3.Text;
			int selectCount = int.Parse(numericUpDown1.Text);

            string disp = "一括入力する内容は" + "\r\n" + "スリーブ：" + selectSleeve + "\r\n" + "不具合原因：" + selectCause + "\r\n" + "数量：" + numericUpDown1.Text + "\r\n" + "で間違いありませんか？";

            DialogResult result = MessageBox.Show(disp, "確認", MessageBoxButtons.YesNo);
			if(result == DialogResult.No)
			{
				return;
			}

			int slCount = 0;
			for(int i = 0; i < listView1.Items.Count; i++)
			{
				if(listView1.Items[i].SubItems[1].Text == selectSleeve)
				{
					int jj = 0;
					for(int k = 0; k < nikuData.Length; k++)
					{
						string resultData = listView1.Items[i].SubItems[15 + jj].Text;
						if(resultData == "OK" || resultData == "-" || resultData == "")
						{
							slCount++;
						}
						jj += 3;
					}
				}
			}
			if(slCount < selectCount)
			{
				MessageBox.Show("一括入力する数量が大きすぎます", "確認", MessageBoxButtons.OK);
				return;
			}


            timer1.Enabled = false;

			//CSV更新:別ファイルに全て読んで、一部を書き換えてファイル名を変える
            string line = "";
            StreamReader reader = null;
            StreamWriter writer = null;
            string path = "";

			List<sameKataList> list = new List<sameKataList>();

			try
			{
                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
                string[] lines = File.ReadAllLines(currentCsvFile);
	            int lineMax = lines.Length;//CSVの行数取得
                path = currentCsvFile + ".tmp";
				writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

				int index = 0;
				while(reader.Peek() >= 0)
				{
					line = reader.ReadLine();

					if(index == 0)
					{
						index++;
						continue;
					}

					string[] cols = line.Split(',');

		            sameKataList katalist = new sameKataList();//とりあえず全部、リストにいれる
		            katalist.sleeveNumber = cols[11];
					katalist.lineNumber = index;
					katalist.result = new string[] {"", "", "", "", "", ""};
					katalist.isEdit = new bool[] {false, false, false, false, false, false };

					int jj = 0;
					for(int k = 0; k < nikuData.Length; k++)
					{
						katalist.result[k] = cols[60 + jj];
						jj += 3;
					}
					list.Add(katalist);

					index++;
				}

				int editCount = 0;
				for(int i = (list.Count - 1); i >= 0; i--)//下から検索でなくてもいいはず
				{
					if(list[i].sleeveNumber == selectSleeve)
					{
						for(int k = 0; k < nikuData.Length; k++)
						{
							if(list[i].result[k] == "OK" || list[i].result[k] == "-" || list[i].result[k] == "")
							{
								sameKataList kataList = list[i];
								kataList.isEdit[k] = true;
								list[i] = kataList;

								editCount++;

								if(editCount == selectCount)//要求の削除数に達した
								{
									break;
								}
							}
						}
						if(editCount == selectCount)//要求の削除数に達した
						{
							break;
						}
					}
				}

				reader.BaseStream.Seek(0, SeekOrigin.Begin);//先頭に戻す
				reader.DiscardBufferedData();


				//NGの為、肉厚値は基準値とする
				double nikuUplimit = double.Parse(numericUpDown12.Text);
				double nikuDnlimit = double.Parse(numericUpDown13.Text);
				double initData = (nikuUplimit + nikuDnlimit) / 2;
				string initNiku = initData.ToString("F3");
				
				int pos = 0;
				while(reader.Peek() >= 0)
				{
					line = reader.ReadLine();
					if(pos == 0)
					{
						writer.WriteLine(line);
						pos++;
						continue;
					}

					if(pos == list[pos - 1].lineNumber)
					{
						string[] cols = line.Split(',');
						string buf = "";
						int jj = 0;
						for(int i = 0; i < cols.Length; i++)
						{
							if(i == 0)
							{
								buf = cols[i];
							}
							else if(i == 59 || i == 62 || i == 65 || i == 68 || i == 71 || i == 74)
							{
								if(list[pos - 1].isEdit[jj])//一括NG入力した箇所のみ書き換える
								{
									buf += "," + initNiku;
								}
								else
								{
									buf += "," + cols[i];
								}
							}
							else if(i == 60 || i == 63 || i == 66 || i == 69 || i == 72 || i == 75)
							{
								if(list[pos - 1].isEdit[jj])//一括NG入力した箇所のみ書き換える
								{
									buf += "," + "NG";
								}
								else
								{
									buf += "," + cols[i];
								}
							}
							else if(i == 61 || i == 64 || i == 67 || i == 70 || i == 73 || i == 76)
							{
								if(list[pos - 1].isEdit[jj])//一括NG入力した箇所のみ書き換える
								{
									buf += "," + selectCause;
								}
								else
								{
									buf += "," + cols[i];
								}
								jj++;
							}
							else
							{
								buf += "," + cols[i];
							}
						}
						writer.WriteLine(buf);
					}
					else
					{
						writer.WriteLine(line);
					}

					pos++;
				}
				
				//ListViewも更新
				int count = 0;
				for(int i = 0; i < listView1.Items.Count; i++)
				{
					listView1.Items[i].UseItemStyleForSubItems = false;
					if(listView1.Items[i].SubItems[1].Text == selectSleeve)
					{
						int jj = 0;
						for(int k = 0; k < nikuData.Length; k++)
						{
							if(list[list.Count - 1 - i].isEdit[k])
							{
								listView1.Items[i].SubItems[14 + jj].Text = initNiku;
								listView1.Items[i].SubItems[15 + jj].Text = "NG";
								listView1.Items[i].SubItems[16 + jj].Text = selectCause;
								listView1.Items[i].SubItems[14 + jj].BackColor = Color.Red;
								listView1.Items[i].SubItems[15 + jj].BackColor = Color.Red;
								listView1.Items[i].SubItems[16 + jj].BackColor = Color.Red;
								count++;
							}
                            jj += 3;
                        }
                    }
				}
			
                //ヘッダの幅を自動調節
                listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
			}
			catch (System.IO.IOException ex)
			{
				string errorStr = "締めの更新時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			finally
			{
				if(reader != null)
				{
					reader.Close();
					//元ファイル削除
					File.Delete(@currentCsvFile);
				}
				if(writer != null)
				{
					writer.Close();
					//一時ファイル→元ファイルへファイル名変更
					System.IO.File.Move(@path, @currentCsvFile);
				}

			}

            timer1.Enabled = true;


			comboBox1.Text = "";
			comboBox3.Text = "";
			numericUpDown1.Text = "1";
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            timer2.Enabled = true;
            currentKaatsujikanUp = (int)(numericUpDown3.Value);

			for(int i = 0; i < SETDATA.seihinName.Length; i++)
			{
				if(SETDATA.seihinName[i] == currentKaatsujikan)
				{
					SETDATA.KaatsuJikanUpper[i] = currentKaatsujikanUp.ToString();
					SETDATA.KaatsuJikanLower[i] = currentKaatsujikanLo.ToString();
					break;
				}
			}
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            timer2.Enabled = true;
			currentKaatsujikanLo = (int)(numericUpDown4.Value);

			for(int i = 0; i < SETDATA.seihinName.Length; i++)
			{
				if(SETDATA.seihinName[i] == currentKaatsujikan)
				{
					SETDATA.KaatsuJikanUpper[i] = currentKaatsujikanUp.ToString();
					SETDATA.KaatsuJikanLower[i] = currentKaatsujikanLo.ToString();
					break;
				}
			}
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            WriteDataToXml();
        }

        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {
            timer3.Enabled = false;
            timer3.Enabled = true;
			nikuUpLimit = numericUpDown12.Value.ToString();

			for(int i = 0; i < SETDATA.nikuUpper.Length; i++)
			{
				if(SETDATA.seihinName[i] == currentSeihin)
				{
					SETDATA.nikuUpper[i] = nikuUpLimit;
					SETDATA.nikuLower[i] = nikuLoLimit;
					break;
				}
			}
			WriteDataToXml();
        }

        private void numericUpDown13_ValueChanged(object sender, EventArgs e)
        {
            timer3.Enabled = false;
            timer3.Enabled = true;
			nikuLoLimit = numericUpDown13.Value.ToString();

			for(int i = 0; i < SETDATA.nikuUpper.Length; i++)
			{
				if(SETDATA.seihinName[i] == currentSeihin)
				{
					SETDATA.nikuUpper[i] = nikuUpLimit;
					SETDATA.nikuLower[i] = nikuLoLimit;
					break;
				}
			}
			WriteDataToXml();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false;
			WriteDataToXml();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
			SETDATA.selectedMachine = comboBox2.Text;
			WriteDataToXml();
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
			SETDATA.selectedOperator = comboBox12.Text;
			WriteDataToXml();
        }

        private void button4_Click(object sender, EventArgs e)
        {
			if(setform == null || setform.IsDisposed)
			{
				setform = new SettingForm();
				setform.ShowDialog();

                setform.Dispose();
                setform = null;

                ReadDataFromXml();

				//今成型している機種であれば画面にも反映させる
				int index = 0xFFFF;
				for(int i = 0; i < SETDATA.seihinName.Length; i++)
				{
					if(SETDATA.seihinName[i] == label1.Text)
					{
						if(SETDATA.KaatsuJikanUpper[i] != "" && SETDATA.KaatsuJikanLower[i] != "")
						{
							index = i;
							break;
						}
					}
				}
				if(index != 0xFFFF)
				{
					currentKaatsujikanUp = int.Parse(SETDATA.KaatsuJikanUpper[index]);
					currentKaatsujikanLo = int.Parse(SETDATA.KaatsuJikanLower[index]);
					numericUpDown3.Text = currentKaatsujikanUp.ToString();
					numericUpDown4.Text = currentKaatsujikanLo.ToString();

					numericUpDown12.Text = Form1.SETDATA.nikuUpper[index];
					numericUpDown13.Text = Form1.SETDATA.nikuLower[index];
				}
			}
        }
    }
}
