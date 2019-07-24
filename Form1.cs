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
    public partial class Form1 : Form
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
			public string result;
			public int lineNumber;
			public bool isEdit;
		}

        static public SYSSET SETDATA = new SYSSET();

		public string InfoStr = "";
		public int totalSleeve = 0;
		public int currentSleeve = 0;
		public int seikeiCount;
		public int updateMode = 0;

		public string currentOperator = "";
		public string currentHousharitsu = "";

		public string shimeSign = "";
		public string sleeveNo = "";
		public string z3Value = "";
		public string ct1Value = "";
		public string ct2Value = "";
		public string cc1Value = "";
		public string cc2Value = "";
		public string cc3Value = "";
		public string cc32Value = "";
		public string cpValue = "";
		public string tactValue = "";
		public string Tkeisu = "";
		public string z3hosei = "";
		public string option1 = "";
		public string option2 = "";
		public string option3 = "";
		public string timeValue = "";
		public string dateValue = "";

		public string nikuEdit = "";
		public string nikuUpLimit = "";
		public string nikuLoLimit = "";
		public string nikuData = "";
		public string nikuResult = "";
		public string resultCause = "";

		public string backLslFile = "";
		public string backTimeValue = "";
		
		public string dmyseikei = "";
		public string dmyshot = "";

		public string kataNo = "";
		public string shotCnt = "";
		public string shukaiCnt = "";
		
		public DateTime backFileTime;
		public string currentCsvFile = "";
		public StreamWriter logWriter;
		
		public string currentSeihin = "";
		public string currentMachine = "";
		public string backSeihinName = "";
		public string commandDate = "";
		public string previousTime = "";

		public string currentKaatsujikan = "";
		public int currentKaatsujikanUp = 0;
		public int currentKaatsujikanLo = 0;

		public int hidden = 1;

		FormCalender formCalender = null;
		SettingForm setform = null;

        public Form1()
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
//            columnCC1Value = new ColumnHeader();
//            columnCC2Value = new ColumnHeader();
//            columnCC3Value = new ColumnHeader();
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
            columnNikuatsuUpper.Text = "肉厚上限";
            columnNikuatsuValue.Text = "測定値";
            columnNikuatsuLower.Text = "肉厚下限";
            columnResult.Text = "合否";
            columnNgCause.Text = "原因";
            columnShotCount.Text = "ｼｮｯﾄ数";
            columnSeikeiCount.Text = "成型数";

            ColumnHeader[] colHeaderRegValue =
              { columnShime, columnSleeve, columnTkeisu, columnInitialTemp, columnSeikeiTemp, columnKaatsuTime, columnZ3Value, columnZ3hosei, /*columnCC1Value, columnCC2Value, columnCC3Value, */
              columnCpValue, columnTact, columnDate, columnTimeStamp, columnNikuatsuUpper, columnNikuatsuValue, columnNikuatsuLower, columnResult, columnNgCause, columnShotCount, columnSeikeiCount};
            listView1.Columns.AddRange(colHeaderRegValue);

			//ヘッダの幅を自動調節
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

			listView1.HideSelection = false;

			string searchStr = "";
			if(SETDATA.machineType == 0)
			{
				searchStr = "LS";
			}
			else if(SETDATA.machineType == 1)
			{
				searchStr = "NQD";
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
			if(SETDATA.machineType == 0)
			{
				existFile = "*_LS*号機.csv";
			}
			else if(SETDATA.machineType == 1)
			{
				existFile = "*_NQD*号機.csv";
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
										backLslFile = fields[j];
									}
								}
								else if(j == 20)//time
								{
									timeValue = fields[j];
									if(timeValue.Length == 19)
									{
										timeValue = fields[j].Substring(11, 8);
										dateValue = fields[j].Substring(0, 10);

										commandDate = dateValue;
									}
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
								}
								else if(j == 9)//ct2Value
								{
									ct2Value = fields[j];
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

					string[] item1 = {shimeSign, sleeveNo, Tkeisu, ct1Value, ct2Value, cc32Value, z3Value, z3hosei, 
					cpValue, tactValue, dateValue, timeValue, 
					nikuUpLimit, nikuData, nikuLoLimit, nikuResult, resultCause, shotCnt, shukaiCnt};

					listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
		            listView1.Font = new System.Drawing.Font("Times New Roman", 10, System.Drawing.FontStyle.Regular);

					if(nikuResult == "OK")
					{
			            listView1.Items[0].BackColor = Color.Lime;
					}
					else if(nikuResult == "NG")
					{
			            listView1.Items[0].BackColor = Color.Red;
					}

					if(sleeveNo == "0" || sleeveNo == "" || sleeveNo == "D" || sleeveNo == "SD" || sleeveNo == "先")
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
					string kaatsuTimeStr = tmp_cc32Value.Substring(0, (tmp_cc32Value.Length - 1));
					int kaatsuTime = int.Parse(kaatsuTimeStr);

					string slNum = listView1.Items[i].SubItems[1].Text;
					if(slNum != "0" && slNum != "" && slNum != "D" && slNum != "SD" && slNum != "先")
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
			if(strResults[19].IndexOf(":") > 0)//正常：時刻表示
			{
				string tmptimeValue = strResults[19].Substring(0, 8);//不要な"を削除
				timeValue = strResults[19] = tmptimeValue;
			}
			else
			{
				if(strResults[19] == "k=")
				{
					if(strResults[21] == "C=")
					{
						string tmptimeValue = strResults[23].Substring(0, 8);//不要な"を削除
						timeValue = strResults[23] = tmptimeValue;
					}
					else
					{
						string tmptimeValue = strResults[21].Substring(0, 8);//不要な"を削除
						timeValue = strResults[21] = tmptimeValue;
					}
				}
				else if(strResults[19] == "C=")
				{
					string tmptimeValue = strResults[21].Substring(0, 8);//不要な"を削除
					timeValue = strResults[21] = tmptimeValue;
				}


				if(strResults[20] == "k=")
				{
					if(strResults[22] == "C=")
					{
						string tmptimeValue = strResults[24].Substring(0, 8);//不要な"を削除
						timeValue = strResults[24] = tmptimeValue;
					}
					else
					{
						string tmptimeValue = strResults[22].Substring(0, 8);//不要な"を削除
						timeValue = strResults[22] = tmptimeValue;
					}
				}
				else if(strResults[20] == "C=")
				{
					string tmptimeValue = strResults[22].Substring(0, 8);//不要な"を削除
					timeValue = strResults[22] = tmptimeValue;
				}
			}
		}

        private void timer1_Tick(object sender, EventArgs e)
        {
			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();

			int boundPos = stCurrentDir.LastIndexOf("\\");
            string str1 = stCurrentDir.Substring(0, boundPos);
            str1 += "\\Data";
            stCurrentDir = str1;

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(stCurrentDir);
            IEnumerable<System.IO.FileInfo> files = di.EnumerateFiles("*.lsl", System.IO.SearchOption.TopDirectoryOnly);

			string strTime = "2019/01/01 01:23:45";
			DateTime lastDt = DateTime.Parse(strTime);
			DateTime destDt;
			string lastFileName = "";

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
				string errorStr = "他のアプリがLSLファイルを開いている可能性があります";
			    System.Console.WriteLine(errorStr);
			    System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			    return;
			}

			string str_back = backFileTime.ToString();
			string str_last = lastDt.ToString();
            if (backLslFile == lastFileName && str_back == str_last)//最新ファイル名が同一かつ更新日時が同一の場合
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
				if(backLslFile != "")
				{
					lastFileName = fileList[k];
					lastDt = dateList[k];
					if(!isExist && lastFileName != backLslFile)
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
					//最新のLSLファイルをWORK用一時ファイルにコピーし、以降は一時ファイルを扱う。
					//最新LSLファイルは成型機VBアプリもアクセスするため、競合区間を最小限にするため。
					File.Copy(@srcFile, @dstFile);
		            fp = new FileStream(dstFile, FileMode.Open, FileAccess.Read);
		            sr = new StreamReader(fp, System.Text.Encoding.GetEncoding("Shift_JIS"));
				}
				catch (System.IO.IOException ex)
				{
					string errorStr = "LSLファイルコピー失敗またはLSLファイルが開けなかった可能性があります";
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
				int endOfPos = 0;
	            string line = "";
				bool isBoundary = false;
				bool isSpecial = false;
				int headerPos = 0;
	            while (sr.EndOfStream == false)
	            {
	                line = sr.ReadLine();

					//先頭行のParse
					if(pos == 0)
					{
	                    string[] stringValues = line.Split(new char[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
	                    currentSeihin = stringValues[0].Substring(1, 7);
						if(commandDate == "")
						{
							commandDate = stringValues[1].Substring(0, 10);
						}

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

	                    pos++;
	                    continue;
	                }

					string[] Linestring = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
					if(Linestring[0] == "\"" && Linestring[1] == "No.")//ヘッダの位置記憶
					{
						headerPos = pos;
					}

	                //コマンドとの境界文字列が含まれている場合
	                if (line.Contains("＊＊＊＊＊"))//lslファイルにより数が異なるがとりあえず5つ
	                {
						if(pos == 2)
						{
							isSpecial = true;
						}
						else
						{
							endOfPos = pos - 1;//境界文字列の1行前を記憶しておく
							isBoundary = true;
						}
						pos++;
						break;
					}

	                pos++;
				}

				if(pos == 2 || isSpecial)//異常系は抜ける：ファイル名＋ヘッダ、ファイル名＋ヘッダ＋＊＊＊
				{
					//一時ファイルを閉じる
					if(sr != null)
					{
						sr.Close();
					}
					File.Delete(@dstFile);//一時ファイルは削除

					continue;
				}

				//最新ファイル名が直前のファイル名と異なる場合、ファイル名と更新時間を入れ替え
				if(backLslFile != lastFileName)
				{
					backLslFile = lastFileName;
					backFileTime = lastDt;
					updateMode = 2;
				}
				//最新ファイル名が直前のファイル名と同じで時間が異なる場合、更新時間を入れ替え
				str_back = backFileTime.ToString();
				str_last = lastDt.ToString();
				if(backLslFile == lastFileName && str_back != str_last)
				{
					backFileTime = lastDt;
					updateMode = 1;
				}

				//境界文字列が含まれていた場合、1行前がスリーブの最新値の為、再度読み込む
				if(isBoundary)
				{
					pos = 0;
					sr.DiscardBufferedData();//一度先頭に戻す
					sr.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
		            while (sr.EndOfStream == false)
		            {
	                    line = sr.ReadLine();
	                    if (pos == endOfPos)
						{
							break;
						}
						pos++;
					}
				}

	            //最後行のParse
	            string lastline = line;
	            //文字列を各項目に分割
	            string[] strResults = lastline.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

	            //コンマがあれば除く
				for(int i = 0; i < strResults.Length; i++)
				{
	                strResults[i] = strResults[i].Replace(",", "");
	            }

	            //内部変数にコピー
	            int sleeveLen = strResults[0].Length;
	            string tmpSleeveNo = strResults[0].Substring(1, (sleeveLen - 1));//不要な"を削除
	            strResults[0] = tmpSleeveNo;

				//前回からの更新がいくつあるか確認し、その差分を全て盛り込む
				int lastPos = 0;//直前のログ(listView上)
				int justPos = 0;//最新のログ

				if(isBoundary)
				{
	                justPos = pos;//行末に＊＊＊がある為、直前の行(最新ログ)で抜ける
				}
				else
				{
	               justPos = pos - 1;//行末に＊＊＊がない為、行末の行(最新ログ)まで読み込む
				}

				if(updateMode == 1)
				{
					sr.DiscardBufferedData();//一度先頭に戻す
					sr.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
		            while (sr.EndOfStream == false)
					{
						//ログファイル内の時間を取得する
						lastline = sr.ReadLine();
			            if(lastPos < 2)
			            {
							lastPos++;
							continue;
						}

	//					if(isBoundary && lastPos == endOfPos)//同一ログファイルが(＊なし→＊あり)に変わった時、はじく
	//					{
	//                        lastPos = 1;
	//                        break;
	//					}
			            
			            strResults = lastline.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
						if(SETDATA.machineType == 0)//LS
						{
							string tmptimeValue = strResults[16].Substring(0, 8);//不要な"を削除
							timeValue = strResults[16] = tmptimeValue;
						}
						else if(SETDATA.machineType == 1)//NQD
						{
							if(strResults[18].IndexOf(":") > 0)//正常：時刻表示
							{
								string tmptimeValue = strResults[18].Substring(0, 8);//不要な"を削除
								timeValue = strResults[18] = tmptimeValue;
							}
							else
							{
								ParseLogString(strResults);
							}
						}

						if(listView1.Items.Count > 0)//ListViewに1つでも登録がある
						{
							//リストにある最新のものとログファイルの時間を比較し、差分がいくつあるか数える
							if(timeValue == listView1.Items[0].SubItems[11].Text)
							{
		                        break;
							}
						}
						lastPos++;
						
	//					if(lastPos > justPos)//同一ログファイルが(＊あり→＊なし)に変わり、listViewに一致する時間のログが無い為、下記のループで新ログファイルを全て読む。
	//					{
	//						if(!isBoundary)
	//						{
	//							lastPos = 1;
	//						}
	//					}
					}
				}

				sr.DiscardBufferedData();//一度先頭に戻す
				sr.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
				pos = 0;
	            while(sr.EndOfStream == false)
				{
					lastline = sr.ReadLine();
					if(updateMode == 1)//ログファイルが変わらず更新日時が変わった時　→　更新分のみ読み込む
					{
						if(lastPos >= pos)//ファイル名、ヘッダ部読み飛ばし
						{
							pos++;
							continue;
						}
					}
					else if(updateMode == 2)//ログファイル名が変わっている時　→　新ログファイルから全て読み込む
					{
						if(pos < 2)//ファイル名、ヘッダ部読み飛ばし
						{
							pos++;
							continue;
						}
					}

					if(isBoundary && pos > endOfPos)//終端の＊＊＊に達した時：既存CSVを読込後もログに変更がない時
					{
						continue;
					}

		            strResults = lastline.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

		            //コンマがあれば除く
					for(int i = 0; i < strResults.Length; i++)
					{
		                strResults[i] = strResults[i].Replace(",", "");
		            }

					//変数に代入ここから
					sleeveNo = strResults[0];

					if(SETDATA.machineType == 0)//LS
					{
						kataNo = strResults[1];
						shotCnt = strResults[2];

						z3Value = strResults[3];
						ct1Value = strResults[4];
						ct2Value = strResults[5];
						cc1Value = strResults[6];
						cc2Value = strResults[7];
						cc3Value = strResults[8];
						cc32Value = strResults[9];

						cpValue = strResults[10];
						tactValue = strResults[11];

						string tmpTkeisu = string.Format("{0:#.###}", strResults[12]);
						string tmpz3hosei = string.Format("{0:#.###}", strResults[13]);
						string tmpoption1 = string.Format("{0:000}", strResults[14]);
						string tmpoption2 = string.Format("{0:0000}", strResults[15]);

						Tkeisu = strResults[12] =  tmpTkeisu;
						z3hosei = strResults[13] = tmpz3hosei;
						option1 = strResults[14] = tmpoption1;
						option2 = strResults[15] = tmpoption2;

						string tmptimeValue = strResults[16].Substring(0, 8);//不要な"を削除
						timeValue = strResults[16] = tmptimeValue;

						sleeveNo = kataNo;
					}
					else if(SETDATA.machineType == 1)//NQD
					{
						dmyseikei = strResults[1];
						dmyshot = strResults[2];

						kataNo = strResults[3];
						shotCnt = strResults[4];

						z3Value = strResults[5];
						ct1Value = strResults[6];
						ct2Value = strResults[7];
						cc1Value = strResults[8];
						cc2Value = strResults[9];
						cc3Value = strResults[10];
						cc32Value = strResults[11];
						cpValue = strResults[12];
						tactValue = strResults[13];

						string tmpTkeisu = string.Format("{0:#.###}", strResults[14]);
						string tmpz3hosei = string.Format("{0:#.###}", strResults[15]);
						string tmpoption1 = string.Format("{0:000}", strResults[16]);
						string tmpoption2 = string.Format("{0:0000}", strResults[17]);

						Tkeisu = strResults[14] =  tmpTkeisu;
						z3hosei = strResults[15] = tmpz3hosei;
						option1 = strResults[16] = tmpoption1;
						option2 = strResults[17] = tmpoption2;

						if(strResults[18].IndexOf(":") > 0)//正常：時刻表示
						{
							string tmptimeValue = strResults[18].Substring(0, 8);//不要な"を削除
							timeValue = strResults[18] = tmptimeValue;
						}
						else
						{
							ParseLogString(strResults);
						}

						sleeveNo = kataNo;
		                shukaiCnt = dmyseikei;
					}

					shimeSign = "";
					string yearStr = commandDate.Substring(0, 4);
					string monthStr = commandDate.Substring(5, 2);
					string dayStr = commandDate.Substring(8, 2);
					string executeDate = yearStr + "/" + monthStr + "/" + dayStr;

					if(previousTime != "")//初回以外
					{
                        DateTime pre = DateTime.Parse(previousTime);
                        DateTime tim = DateTime.Parse(timeValue);
                        if (pre > tim)//日付が変わった時
						{
							DateTime ddtt = DateTime.Parse(executeDate);
							ddtt = ddtt.AddDays(1);
							executeDate = ddtt.ToString("yyyy/MM/dd");
							commandDate = ddtt.ToString("yyyy-MM-dd");
						}
					}
					previousTime = timeValue;
					dateValue = executeDate;
					string out_date = dateValue + " " + timeValue;

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
					currentMachine = seikeikiName;

					string logStr = "";
					logStr += "," + seikeikiName;
					logStr += "," + currentSeihin;
					logStr += "," + lastFileName;
					logStr += "," + shimeSign;
					logStr += "," + sleeveNo;
					logStr += "," + z3Value;
					logStr += "," + ct1Value;
					logStr += "," + ct2Value;
					logStr += "," + cc1Value;
					logStr += "," + cc2Value;
					logStr += "," + cc3Value;
					logStr += "," + cc32Value;
					logStr += "," + cpValue;
					logStr += "," + tactValue;
					logStr += "," + Tkeisu;
					logStr += "," + z3hosei;
					logStr += "," + option1;
					logStr += "," + option2;
					logStr += "," + out_date;

					//CSVに肉厚、上限、下限、合否、不良原因も設定
					nikuResult = resultCause = "-";
					nikuData = "";
					logStr += string.Format(",{0:F3},{1:F3},{2:F3}",nikuUpLimit, nikuData, nikuLoLimit);
					logStr += "," + nikuResult + "," + resultCause;

					currentOperator = comboBox12.Text;
					currentHousharitsu = numericUpDown2.Text;
					logStr += "," + currentHousharitsu + "," + currentOperator + "," + shotCnt + "," + shukaiCnt;

					if(backTimeValue != timeValue)//ログファイルが更新されていても、最新時間が同じ場合の念の為ガード
					{
					//lslファイルに更新がある場合
						string fileName = lastFileName.Substring(0, (lastFileName.Length - 4));
						//CSVに保存する
						if(WriteDataToCsv(logStr, currentSeihin, currentMachine))
						{
							backTimeValue = timeValue;

							//リストビューを更新
							AddListView();
						}
					}

					if(justPos == pos)
					{
						break;
					}
					
					pos++;
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
					string errorStr = "LSL一時ファイルが削除できなかった可能性があります";
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

			string[] item1 = {shimeSign, sleeveNo, Tkeisu, ct1Value, ct2Value, cc32Value, z3Value, z3hosei, 
			/*cc1Value, cc2Value, cc3Value,*/ cpValue, tactValue, dateValue, timeValue, 
			nikuUpLimit, nikuData, nikuLoLimit, nikuResult, resultCause, shotCnt, shukaiCnt};

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

			if(sleeveNo == "0" || sleeveNo == "" || sleeveNo == "D" || sleeveNo == "SD" || sleeveNo == "先")
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

			string kaatsuTimeStr = cc32Value.Substring(0, (cc32Value.Length - 1));
			int kaatsuTime = int.Parse(kaatsuTimeStr);
			if(sleeveNo != "0" && sleeveNo != "" && sleeveNo != "D" && sleeveNo != "SD" && sleeveNo != "先")
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
            SETDATA.load(ref Form1.SETDATA);
		}

		static public void WriteDataToXml()
		{
            SETDATA.save(Form1.SETDATA);
		}

		public bool WriteDataToCsv(string logStr, string seihin, string machine)
		{
            try
            {
		        // appendをtrueにすると，既存のファイルに追記
		        //         falseにすると，ファイルを新規作成する
		        var append = false;
		        // 出力用のファイルを開く
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = currentCsvFile = stCurrentDir + "\\PgmOut_" + seihin + "_" + machine + ".csv";

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
                        buf += string.Format(",ｽﾘｰﾌﾞNo");
						buf += string.Format(",Z3");
						buf += string.Format(",ct1");
						buf += string.Format(",ct2");
						buf += string.Format(",cc1");
						buf += string.Format(",cc2");
						buf += string.Format(",cc3");
						buf += string.Format(",cc3-2");
						buf += string.Format(",cp");
						buf += string.Format(",ﾀｸﾄ");
						buf += string.Format(",T係数");
						buf += string.Format(",Z3補正");
						buf += string.Format(",ｵﾌﾟｼｮﾝ1");
						buf += string.Format(",ｵﾌﾟｼｮﾝ2");
						buf += string.Format(",ﾀｲﾑｽﾀﾝﾌﾟ");
						buf += string.Format(",肉厚上限");
						buf += string.Format(",肉厚測定値");
						buf += string.Format(",肉厚下限");
						buf += string.Format(",合否");
						buf += string.Format(",不良原因");
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


			public int machineType;//0:LS, 1:NQD
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

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
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

				if(listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "" || 
				listView1.Items[index].SubItems[1].Text == "D" || listView1.Items[index].SubItems[1].Text == "SD" || 
				listView1.Items[index].SubItems[1].Text == "先")
				{
					timer1.Enabled = true;
					return;
				}

				string justSleeve = listView1.Items[index].SubItems[1].Text;
				nikuUpLimit = listView1.Items[index].SubItems[12].Text;//肉厚上限
				nikuLoLimit = listView1.Items[index].SubItems[14].Text;//肉厚下限
				nikuEdit = listView1.Items[index].SubItems[13].Text;//肉厚入力値
				resultCause = listView1.Items[index].SubItems[16].Text;//不良原因
				string currentShotCount = listView1.Items[index].SubItems[17].Text;//ショット数

                //登録・編集画面の表示
                Form5 form5 = new Form5();
				form5.SetInfo(justSleeve, currentShotCount, nikuUpLimit, nikuLoLimit, nikuEdit, resultCause);
	            form5.ShowDialog();

	            InfoStr = form5.EditInfo;
	            if(InfoStr == "")
	            {
					timer1.Enabled = true;
					return;
				}

				string[] strline = InfoStr.Split(',');

				double nikuData = double.Parse(strline[0]);
				double nikuUpData = double.Parse(nikuUpLimit);
				double nikuLoData = double.Parse(nikuLoLimit);
				bool isOK = true;

				if(strline[1] != "なし")//不良原因
				{
					isOK = false;
				}
				else
				{
					if(nikuData < nikuLoData || nikuUpData < nikuData)
					{
						isOK = false;
					}
				}

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
								else if(i == 21)//肉厚上限
								{
									buf += "," + nikuUpLimit;
								}
								else if(i == 22)//肉厚値
								{
									buf += "," + strline[0];
								}
								else if(i == 23)//肉厚下限
								{
									buf += "," + nikuLoLimit;
								}
								else if(i == 24)//合否
								{
									if(isOK)
									{
										buf += "," + "OK";
									}
									else
									{
										buf += "," + "NG";
									}
								}
								else if(i == 25)//不良原因
								{
									if(isOK)
									{
										buf += ",-";
									}
									else
									{
										buf += "," + strline[1];
									}
								}
								else if(i == 28)//ショット数
								{
									buf += "," + strline[2];
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
					listView1.Items[index].SubItems[13].Text = string.Format("{0:#.###}", strline[0]);
					listView1.Items[index].SubItems[12].Text = nikuUpLimit;//肉厚上限
					listView1.Items[index].SubItems[14].Text = nikuLoLimit;//肉厚下限
					listView1.Items[index].SubItems[17].Text = strline[2];//ショット数
					if(isOK)
					{
						listView1.Items[index].SubItems[15].Text = "OK";
						listView1.Items[index].SubItems[16].Text = "-";
						listView1.Items[index].BackColor = Color.Lime;
                    }
					else
					{
						listView1.Items[index].SubItems[15].Text = "NG";
						listView1.Items[index].SubItems[16].Text = strline[1];
						listView1.Items[index].BackColor = Color.Red;
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

        private void Form1_Shown(object sender, EventArgs e)
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
			if(hidden == 0)
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
			string mes = "アプリを終了すると成型機からのデータを受け取れなくなります！" + "\r\n" + "本当に終了しますか？";
			DialogResult result = MessageBox.Show(mes, "PGM成型機監視アプリ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
			if(result == DialogResult.Yes)
			{
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
                string result = "";
                string ngcause = "";
                string shotCount = "";
                string seikeiCount = "";
				string nikuatsuSokutei = "";
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
                    else if(j == 13)//肉厚測定値
					{
						nikuatsuSokutei = listitem;
					}
                    else if(j == 15)//合否
					{
						result = listitem;
					}
                    else if(j == 16)//不良原因
					{
						ngcause = listitem;
					}
                    else if(j == 1)//スリーブ番号
					{
						sleeveNumber = listitem;
					}
                    else if(j == 17)//ショット数
					{
						shotCount = listitem;
					}
                    else if(j == 18)//成型数
					{
						seikeiCount = listitem;
					}

                }

                listDT = listDate + " " + listTime;
                DateTime dttarget = DateTime.Parse(listDT);
                if(dtsta <= dttarget && dttarget <= dtend)
                {
					allSleeve++;
					if(result == "OK")
					{
	                    okCount++;
	                }
					if(result == "NG")
					{
	                    ngCount++;

						if(ngcause == "肉厚不良")
						{
							nikuatsuNG++;//肉厚不良
						}
						else if(ngcause == "キズ")
						{
							kizuNG++;//キズ
						}
						else if(ngcause == "ブツ")
						{
							butsuNG++;//ブツ
						}
						else if(ngcause == "ヤケ")
						{
							yakeNG++;//ヤケ
						}
						else if(ngcause == "ヒビ/クラック")
						{
							hibicrackNG++;//ヒビ/クラック
						}
						else if(ngcause == "ガスキズ")
						{
							gasukizuNG++;//ガスキズ
						}
						else if(ngcause == "放射キズ")
						{
							houshakizuNG++;//放射キズ
						}
						else if(ngcause == "ギラツキ/クモリ")
						{
							giratsukikumoriNG++;//ギラツキ/クモリ
						}
						else if(ngcause == "偏肉/面ダレ")
						{
							hennikumendareNG++;//偏肉/面ダレ
						}
						else if(ngcause == "汚れ")
						{
							yogoreNG++;//汚れ
						}
						else if(ngcause == "ほこり")
						{
							hokoriNG++;//ほこり
						}
						else if(ngcause == "形状精度")
						{
							keijoseidoNG++;//形状精度
						}
						else if(ngcause == "その他")
						{
							etcNG++;//その他
						}
	                }
	                
	                if(nikuatsuSokutei != "-" && nikuatsuSokutei != "")
					{
		                nikuData = double.Parse(nikuatsuSokutei.ToString());
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

					if(sleeveNumber != "0" && sleeveNumber != "" && sleeveNumber != "D" && sleeveNumber != "SD" && sleeveNumber != "先")
					{
						workSleeve++;

						//未登録なら追加
	                    if(list.Find(m => m.sleeveNumber == sleeveNumber).sleeveNumber != sleeveNumber)
		                {
	                        SleeveList sl = new SleeveList();
	                        sl.sleeveNumber= sleeveNumber;
							sl.shotCount = int.Parse(shotCount);

							if(result == "OK")
							{
								sl.ryohinSuu++;
							}
							else if(result == "NG")
							{
								sl.furyouSuu++;
							}

							//SL毎の不良内訳
							if(ngcause == "肉厚不良")
							{
								sl.nikuatsuNG++;
							}
							else if(ngcause == "キズ")
							{
								sl.kizuNG++;
							}
							else if(ngcause == "ブツ")
							{
								sl.butsuNG++;
							}
							else if(ngcause == "ヤケ")
							{
								sl.yakeNG++;
							}
							else if(ngcause == "ヒビ/クラック")
							{
								sl.hibicrackNG++;
							}
							else if(ngcause == "ガスキズ")
							{
								sl.gasukizuNG++;
							}
							else if(ngcause == "放射キズ")
							{
								sl.houshakizuNG++;
							}
							else if(ngcause == "ギラツキ/クモリ")
							{
								sl.giratsukikumoriNG++;
							}
							else if(ngcause == "偏肉/面ダレ")
							{
								sl.hennikumendareNG++;
							}
							else if(ngcause == "汚れ")
							{
								sl.yogoreNG++;
							}
							else if(ngcause == "ほこり")
							{
								sl.hokoriNG++;
							}
							else if(ngcause == "形状精度")
							{
								sl.keijoseidoNG++;
							}
							else if(ngcause == "その他")
							{
								sl.etcNG++;
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

									if(result == "OK")
									{
										causeList.ryohinSuu++;
									}
									else if(result == "NG")
									{
										causeList.furyouSuu++;
									}

									//SL毎の不良内訳
									if(ngcause == "肉厚不良")
									{
										causeList.nikuatsuNG++;
									}
									else if(ngcause == "キズ")
									{
										causeList.kizuNG++;
									}
									else if(ngcause == "ブツ")
									{
										causeList.butsuNG++;
									}
									else if(ngcause == "ヤケ")
									{
										causeList.yakeNG++;
									}
									else if(ngcause == "ヒビ/クラック")
									{
										causeList.hibicrackNG++;
									}
									else if(ngcause == "ガスキズ")
									{
										causeList.gasukizuNG++;
									}
									else if(ngcause == "放射キズ")
									{
										causeList.houshakizuNG++;
									}
									else if(ngcause == "ギラツキ/クモリ")
									{
										causeList.giratsukikumoriNG++;
									}
									else if(ngcause == "偏肉/面ダレ")
									{
										causeList.hennikumendareNG++;
									}
									else if(ngcause == "汚れ")
									{
										causeList.yogoreNG++;
									}
									else if(ngcause == "ほこり")
									{
										causeList.hokoriNG++;
									}
									else if(ngcause == "形状精度")
									{
										causeList.keijoseidoNG++;
									}
									else if(ngcause == "その他")
									{
										causeList.etcNG++;
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
        
        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if (listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "" || 
			listView1.Items[index].SubItems[1].Text == "D" || listView1.Items[index].SubItems[1].Text == "SD" || 
			listView1.Items[index].SubItems[1].Text == "先")
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

        private void ToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if (listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "" || 
			listView1.Items[index].SubItems[1].Text == "D" || listView1.Items[index].SubItems[1].Text == "SD" || 
			listView1.Items[index].SubItems[1].Text == "先")
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

        private void ToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if (listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "" || 
			listView1.Items[index].SubItems[1].Text == "D" || listView1.Items[index].SubItems[1].Text == "SD" || 
			listView1.Items[index].SubItems[1].Text == "先")
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

        private void ToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;
            if (listView1.Items[index].SubItems[1].Text == "0" || listView1.Items[index].SubItems[1].Text == "" ||
            listView1.Items[index].SubItems[1].Text == "D" || listView1.Items[index].SubItems[1].Text == "SD" ||
            listView1.Items[index].SubItems[1].Text == "先")
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
            string dailyPath = dailyDir + "\\" + currentSeihin + "_" + currentMachine + "_" + sta_sign + "to" + end_sign + "_" + dailysign + ".csv";

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
                string dailyPath = dailyDir + "\\" + currentSeihin + "_" + currentMachine + "_" + sta_sign + "to" + end_sign + "_" + dailysign + ".csv";

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
						if(sleeveNumber != "0" && sleeveNumber != "" && sleeveNumber != "D" && sleeveNumber != "SD" && sleeveNumber != "先")
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
				if(listView1.Items[i].SubItems[1].Text == selectSleeve &&
				(listView1.Items[i].SubItems[15].Text == "OK" || listView1.Items[i].SubItems[15].Text == "-"))
				{
					slCount++;
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
					string[] cols = line.Split(',');

		            sameKataList katalist = new sameKataList();//とりあえず全部、リストにいれる
		            katalist.sleeveNumber = cols[6];
                    katalist.result = cols[24];
					katalist.lineNumber = index;
					list.Add(katalist);

					index++;
				}

				int editCount = 0;
				for(int i = (list.Count - 1); i >= 0; i--)
				{
					if((list[i].sleeveNumber == selectSleeve) && 
					(list[i].result == "OK" || list[i].result == "-"))
					{
						sameKataList kataList = list[i];
						kataList.isEdit = true;
						list[i] = kataList;

						editCount++;
					}

					if(editCount == selectCount)
					{
						break;
					}
				}

				reader.BaseStream.Seek(0, SeekOrigin.Begin);//先頭に戻す
				reader.DiscardBufferedData();
				
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

					if(pos == list[pos].lineNumber && list[pos].isEdit)
					{
						string[] cols = line.Split(',');
						string buf = "";
						for(int i = 0; i < cols.Length; i++)
						{
							if(i == 0)
							{
								buf = cols[i];
							}
							else if(i == 24)
							{
								buf += "," + "NG";
							}
							else if(i == 25)
							{
								buf += "," + selectCause;
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
					if(listView1.Items[i].SubItems[1].Text == selectSleeve &&
					(listView1.Items[i].SubItems[15].Text == "OK" || listView1.Items[i].SubItems[15].Text == "-"))
					{
						listView1.Items[i].SubItems[15].Text = "NG";
						listView1.Items[i].SubItems[16].Text = selectCause;
						listView1.Items[i].BackColor = Color.Red;
						count++;
					}

					if(count == selectCount)
					{
						break;
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
