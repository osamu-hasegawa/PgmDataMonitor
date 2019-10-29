using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace PgmDataMonitor
{
    public partial class FormCalender : Form
    {
//PCがスリープ状態に入らないようする start
        #region Win32 API
        [FlagsAttribute]
        public enum ExecutionState : uint
        {
            // 関数が失敗した時の戻り値
            Null = 0,
            // スタンバイを抑止(Vista以降は効かない？)
            SystemRequired = 1,
            // 画面OFFを抑止
            DisplayRequired = 2,
            // 効果を永続させる。ほかオプションと併用する。
            Continuous = 0x80000000,
        }

        [DllImport("user32.dll")]
        extern static uint SendInput(
            uint nInputs,   // INPUT 構造体の数(イベント数)
            INPUT[] pInputs,   // INPUT 構造体
            int cbSize     // INPUT 構造体のサイズ
            );

        [StructLayout(LayoutKind.Sequential)]  // アンマネージ DLL 対応用 struct 記述宣言
        struct INPUT
        {
            public int type;  // 0 = INPUT_MOUSE(デフォルト), 1 = INPUT_KEYBOARD
            public MOUSEINPUT mi;
            // Note: struct の場合、デフォルト(パラメータなしの)コンストラクタは、
            //       言語側で定義済みで、フィールドを 0 に初期化する。
        }

        [StructLayout(LayoutKind.Sequential)]  // アンマネージ DLL 対応用 struct 記述宣言
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;  // amount of wheel movement
            public int dwFlags;
            public int time;  // time stamp for the event
            public IntPtr dwExtraInfo;
            // Note: struct の場合、デフォルト(パラメータなしの)コンストラクタは、
            //       言語側で定義済みで、フィールドを 0 に初期化する。
        }

        // dwFlags
        const int MOUSEEVENTF_MOVED = 0x0001;
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;  // 左ボタン Down
        const int MOUSEEVENTF_LEFTUP = 0x0004;  // 左ボタン Up
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;  // 右ボタン Down
        const int MOUSEEVENTF_RIGHTUP = 0x0010;  // 右ボタン Up
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;  // 中ボタン Down
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;  // 中ボタン Up
        const int MOUSEEVENTF_WHEEL = 0x0080;
        const int MOUSEEVENTF_XDOWN = 0x0100;
        const int MOUSEEVENTF_XUP = 0x0200;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        const int screen_length = 0x10000;  // for MOUSEEVENTF_ABSOLUTE
        [DllImport("kernel32.dll")]
        static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);
        #endregion
//PCがスリープ状態に入らないようする end

        Form1 form1 = null;
        Form3 form3 = null;
        Form4 form4 = null;
        Form6 form6 = null;
		int formType = 0;

		bool shimeFlg = false;

		string gouki = "";
		string seihin = "";
		string operatorName = "";
		
        public FormCalender()
        {
            InitializeComponent();

			label1.Text = "監　視　中";

            label18.Text = "";
            label19.Text = "";
            label20.Text = "";
            label22.Text = "";

			label36.Text = "";
			label37.Text = "";
			label38.Text = "";
			label39.Text = "";
			label40.Text = "";
			label41.Text = "";
			label42.Text = "";
			label43.Text = "";
			label44.Text = "";
			label45.Text = "";
			label46.Text = "";
			label47.Text = "";
			label48.Text = "";
			label222.Text = "";

			label50.Text = "0";
			label51.Text = "0";
			label52.Text = "0";
			label53.Text = "0";

			//OSのバージョン情報を取得する
			System.OperatingSystem os = System.Environment.OSVersion;
			//Windows NT系か調べる
			if(os.Platform == PlatformID.Win32NT)
			{
			    //メジャーバージョン番号が6未満がXP以前
			    if(os.Version.Major < 6)
			    {
	                monthCalendar1.Font = new Font("MS UI Gothic", 12, FontStyle.Bold);//XP以前だと設定Fontが反映される
	            }
	        }

//			this.MaximizeBox = false;
//			this.MinimizeBox = false;
			this.ControlBox = !this.ControlBox;
            
            timer1.Enabled = true;
            alive_timer.Enabled = true;
        }

        public void SetParentForm(Form form, int type)
        {
			formType = type;
			switch(formType)
			{
			case 0://LS single
			case 1://NQD single
	            form1 = (Form1)form;
				break;
			case 2://HS single
	            form3 = (Form3)form;
				break;
			case 3://LS multi cav
			case 4://NQD multi cav
	            form4 = (Form4)form;
				break;
			case 5://HS multi cav
	            form6 = (Form6)form;
				break;
			default:
				break;
			}
        }

		public void SetSeikeiInfo(DateTime start, DateTime end)
		{
			if(shimeFlg)
			{
				string staDate = "";
				string endDate = "";
	            if(formType == 0 || formType == 1)//LS,NQD
				{
		            form1.GetShimeHani(start, end, ref staDate, ref endDate);

                    if (staDate == "" || endDate == "")
                    {
                        //締めの開始または終了が指定されていない。
                        label49.Text = "締め開始/終了が未指定";
                        return;
                    }

                    start = DateTime.Parse(staDate);
	                end = DateTime.Parse(endDate);
		        }
				else if(formType == 2)//HS
				{
		            form3.GetShimeHani(start, end, ref staDate, ref endDate);

                    if (staDate == "" || endDate == "")
                    {
                        //締めの開始または終了が指定されていない。
                        label49.Text = "締め開始/終了が未指定";
                        return;
                    }

                    start = DateTime.Parse(staDate);
	                end = DateTime.Parse(endDate);
				}
				else if(formType == 3 || formType == 4)//マルチLS,マルチNQD
				{
		            form4.GetShimeHani(start, end, ref staDate, ref endDate);

                    if (staDate == "" || endDate == "")
                    {
                        //締めの開始または終了が指定されていない。
                        label49.Text = "締め開始/終了が未指定";
                        return;
                    }

                    start = DateTime.Parse(staDate);
	                end = DateTime.Parse(endDate);
				}
				else if(formType == 5)//マルチHS
				{
		            form6.GetShimeHani(start, end, ref staDate, ref endDate);

                    if (staDate == "" || endDate == "")
                    {
                        //締めの開始または終了が指定されていない。
                        label49.Text = "締め開始/終了が未指定";
                        return;
                    }

                    start = DateTime.Parse(staDate);
	                end = DateTime.Parse(endDate);
				}
			}
			else
			{
				if(start == end)
				{
					string ss = "00:00:00";
					start = DateTime.Parse(ss);

					string se = "23:59:59";
					end = DateTime.Parse(se);
				}
			}

            string stDate = start.ToString("yyyy/MM/dd HH:mm:ss");
            string enDate = end.ToString("yyyy/MM/dd HH:mm:ss");

		    DateTime selectStaDate = start;
		    DateTime selectEndDate = end;
            label49.Text = selectStaDate.ToString("    yyyy/MM/dd HH:mm:ss") + " ～" + "\r\n" + selectEndDate.ToString("yyyy/MM/dd HH:mm:ss");

            label18.Text = "0";
            label19.Text = "0";
            label20.Text = "0";
            label22.Text = "0";

			label36.Text = "0";
			label37.Text = "0";
			label38.Text = "0";
			label39.Text = "0";
			label40.Text = "0";
			label41.Text = "0";
			label42.Text = "0";
			label43.Text = "0";
			label44.Text = "0";
			label45.Text = "0";
			label46.Text = "0";
			label47.Text = "0";
			label48.Text = "0";
			label222.Text = "0";

			label50.Text = "0";
			label51.Text = "0";
			label52.Text = "0";
			label53.Text = "0";

			label54.Text = "金型稼働率(使用数/全数)：";

			label55.Text = "";
			label56.Text = "";
			label57.Text = "";
			label58.Text = "";

			label208.Text = "";
			label209.Text = "";
			label210.Text = "";

			Label [] shotLabel = new Label [] {label70, label71, label72, label73, label74, label75, label76, label77, label225, label226};
			Label [] seikeiLabel = new Label [] {label78, label79, label80, label81, label82, label83, label84, label85, label227, label228};
			Label [] goodLabel = new Label [] {label86, label87, label88, label89, label90, label91, label92, label93, label229, label230};
			Label [] furyouLabel = new Label [] {label94, label95, label96, label97, label98, label99, label100, label101, label231, label232};
			Label [] nikufuLabel = new Label [] {label102, label103, label104, label105, label106, label107, label108, label109, label233, label234};
			Label [] kizuLabel = new Label [] {label110, label111, label112, label113, label114, label115, label116, label117, label235, label236};
			Label [] butsuLabel = new Label [] {label118, label119, label120, label121, label122, label123, label124, label125, label237, label238};
			Label [] yakeLabel = new Label [] {label126, label127, label128, label129, label130, label131, label132, label133, label239, label240};
			Label [] hibicrackLabel = new Label [] {label134, label135, label136, label137, label138, label139, label140, label141, label241, label242};
			Label [] gasukizuLabel = new Label [] {label142, label143, label144, label145, label146, label147, label148, label149, label243, label244};
			Label [] houshsaLabel = new Label [] {label150, label151, label152, label153, label154, label155, label156, label157, label245, label246};
			Label [] giratsukiLabel = new Label [] {label158, label159, label160, label161, label162, label163, label164, label165, label247, label248};
			Label [] hennikuLabel = new Label [] {label166, label167, label168, label169, label170, label171, label172, label173, label249, label250};
			Label [] yogoreLabel = new Label [] {label174, label175, label176, label177, label178, label179, label180, label181, label251, label252};
			Label [] hokoriLabel = new Label [] {label182, label183, label184, label185, label186, label187, label188, label189, label253, label254};
			Label [] keijouseidoLabel = new Label [] {label190, label191, label192, label193, label194, label195, label196, label197, label255, label256};
			Label [] sonotaLabel = new Label [] {label198, label199, label200, label201, label202, label203, label204, label205, label257, label258};
			Label [] tachiageLabel = new Label [] {label214, label215, label216, label217, label218, label219, label220, label221, label259, label260};

			Label [] item = new Label [] {label23, label24, label25, label26, label27, label28, label29, label30, label31, label32, label33, label34, label35};
			Label [] valu = new Label [] {label36, label37, label38, label39, label40, label41, label42, label43, label44, label45, label46, label47, label48};

			//初期化
			for(int i = 0; i < shotLabel.Length; i++)
			{
				shotLabel[i].Text = "";
				seikeiLabel[i].Text = "";
				goodLabel[i].Text = "";
				furyouLabel[i].Text = "";
				nikufuLabel[i].Text = "";
				kizuLabel[i].Text = "";
				butsuLabel[i].Text = "";
				yakeLabel[i].Text = "";
				hibicrackLabel[i].Text = "";
				gasukizuLabel[i].Text = "";
				houshsaLabel[i].Text = "";
				giratsukiLabel[i].Text = "";
				hennikuLabel[i].Text = "";
				yogoreLabel[i].Text = "";
				hokoriLabel[i].Text = "";
				keijouseidoLabel[i].Text = "";
				sonotaLabel[i].Text = "";
				tachiageLabel[i].Text = "";
			}

            string okngStr = "";
            string ngcauseStr = "";
            string nikuInfo = "";
            string sleeveInfo = "";
            double kadou = 0;

			string upper = "";
			string lower = "";
			
			int maxShotCount = 0;


			int totalSleeve = 0;
			string nikuatsuNG = "";
			string kizuNG = "";
			string butsuNG = "";
			string yakeNG = ""; 
			string hibicrackNG = "";
			string gasukizuNG = "";
			string houshakizuNG = "";
			string giratsukikumoriNG = "";
			string hennikumendareNG = "";
			string yogoreNG = "";
			string hokoriNG = "";
			string keijoseidoNG = "";
			string etcNG = "";
			string tachiageNG = "";

            if(formType == 0 || formType == 1)//LS,NQD
			{
	            form1.SetDateTimeToAnalize(stDate, enDate, ref okngStr, ref ngcauseStr, ref nikuInfo, ref sleeveInfo, ref kadou, 
										ref totalSleeve, 
										ref  nikuatsuNG, ref  kizuNG, ref  butsuNG, ref  yakeNG, 
										ref  hibicrackNG, ref  gasukizuNG, ref  houshakizuNG, 
										ref  giratsukikumoriNG, ref  hennikumendareNG, ref  yogoreNG, 
										ref  hokoriNG, ref  keijoseidoNG, ref  etcNG, ref tachiageNG);

				form1.GetDisplayData(ref gouki, ref seihin, ref upper, ref lower, ref operatorName);
				string shotword = string.Format("限界ショット数 ( {0} : 赤 / {1} : 黄 ) : ", (Form1.SETDATA.maxShotCount - 20), (Form1.SETDATA.maxShotCount - 40));
                label17.Text = shotword + Form1.SETDATA.maxShotCount;
            	maxShotCount = Form1.SETDATA.maxShotCount;
            }
            else if(formType == 2)//HS
			{
	            form3.SetDateTimeToAnalize(stDate, enDate, ref okngStr, ref ngcauseStr, ref nikuInfo, ref sleeveInfo, ref kadou, 
										ref totalSleeve, 
										ref  nikuatsuNG, ref  kizuNG, ref  butsuNG, ref  yakeNG, 
										ref  hibicrackNG, ref  gasukizuNG, ref  houshakizuNG, 
										ref  giratsukikumoriNG, ref  hennikumendareNG, ref  yogoreNG, 
										ref  hokoriNG, ref  keijoseidoNG, ref  etcNG, ref tachiageNG);

				form3.GetDisplayData(ref gouki, ref seihin, ref upper, ref lower, ref operatorName);
				string shotword = string.Format("限界ショット数 ( {0} : 赤 / {1} : 黄 ) : ", (Form3.SETDATA.maxShotCount - 20), (Form3.SETDATA.maxShotCount - 40));
                label17.Text = shotword + Form3.SETDATA.maxShotCount;
            	maxShotCount = Form3.SETDATA.maxShotCount;
            }
            else if(formType == 3 || formType == 4)//マルチLS,マルチNQD
			{
	            form4.SetDateTimeToAnalize(stDate, enDate, ref okngStr, ref ngcauseStr, ref nikuInfo, ref sleeveInfo, ref kadou, 
										ref totalSleeve, 
										ref  nikuatsuNG, ref  kizuNG, ref  butsuNG, ref  yakeNG, 
										ref  hibicrackNG, ref  gasukizuNG, ref  houshakizuNG, 
										ref  giratsukikumoriNG, ref  hennikumendareNG, ref  yogoreNG, 
										ref  hokoriNG, ref  keijoseidoNG, ref  etcNG, ref tachiageNG);

				form4.GetDisplayData(ref gouki, ref seihin, ref upper, ref lower, ref operatorName);
				string shotword = string.Format("限界ショット数 ( {0} : 赤 / {1} : 黄 ) : ", (Form4.SETDATA.maxShotCount - 20), (Form4.SETDATA.maxShotCount - 40));
                label17.Text = shotword + Form4.SETDATA.maxShotCount;
            	maxShotCount = Form4.SETDATA.maxShotCount;
            }
            else if(formType == 5)//マルチHS
			{
	            form6.SetDateTimeToAnalize(stDate, enDate, ref okngStr, ref ngcauseStr, ref nikuInfo, ref sleeveInfo, ref kadou, 
										ref totalSleeve, 
										ref  nikuatsuNG, ref  kizuNG, ref  butsuNG, ref  yakeNG, 
										ref  hibicrackNG, ref  gasukizuNG, ref  houshakizuNG, 
										ref  giratsukikumoriNG, ref  hennikumendareNG, ref  yogoreNG, 
										ref  hokoriNG, ref  keijoseidoNG, ref  etcNG, ref tachiageNG);

				form6.GetDisplayData(ref gouki, ref seihin, ref upper, ref lower, ref operatorName);
				string shotword = string.Format("限界ショット数 ( {0} : 赤 / {1} : 黄 ) : ", (Form6.SETDATA.maxShotCount - 20), (Form6.SETDATA.maxShotCount - 40));
                label17.Text = shotword + Form6.SETDATA.maxShotCount;
            	maxShotCount = Form6.SETDATA.maxShotCount;
            }

			for(int i = 0; i < item.Length; i++)
			{
				item[i].ForeColor = Color.Black;
				valu[i].ForeColor = Color.Black;
			}

            string[] line = okngStr.Split(',');

            label209.Text = label18.Text = line[0];
	        label210.Text = label19.Text = line[1];
			label209.ForeColor = label18.ForeColor = Color.Lime;
			label210.ForeColor = label19.ForeColor = Color.Red;
			label20.ForeColor = Color.Blue;
            int okCount = int.Parse(line[0]);
            int ngCount = int.Parse(line[1]);

			if(ngCount > 0)
			{
                string[] ngline = ngcauseStr.Split(',');

				for(int i = 0; i < item.Length; i++)
				{
					valu[i].Text = ngline[i];
					if(int.Parse(ngline[i]) > 0)
					{
						item[i].ForeColor = Color.Red;
						valu[i].ForeColor = Color.Red;
					}
				}
		    }

			//２行に渡るlabelの色を統一
			label60.ForeColor = label30.ForeColor;
			label61.ForeColor = label31.ForeColor;
			label63.ForeColor = label34.ForeColor;
			label64.ForeColor = label29.ForeColor;
			label59.ForeColor = label27.ForeColor;
			label65.ForeColor = label23.ForeColor;

			if(nikuInfo != "")
			{
                string[] nikuline = nikuInfo.Split(',');

				label51.Text = nikuline[0];//平均
				label52.Text = nikuline[1];//最大
				label50.Text = nikuline[2];//最小
				label53.Text = nikuline[3];//標準偏差
			}


			Label [] label = new Label[]{label12, label11, label10, label9, label16, label15, label14, label13, label223, label224};
			for(int i = 0; i < label.Length; i++)
			{
				label[i].Text = "";
				label[i].BackColor = Color.White;
			}

			string[] sleeveline;
			if(sleeveInfo != "")
			{
                sleeveline = sleeveInfo.Split(',');
                int j = 0;
				for(int i = 0; i < sleeveline.Length; i+=5)
				{
					label[j].Text = sleeveline[i];
					shotLabel[j].Text = sleeveline[i + 1];
					seikeiLabel[j].Text = sleeveline[i + 2];
					goodLabel[j].Text = sleeveline[i + 3];
					furyouLabel[j].Text = sleeveline[i + 4];

					if((maxShotCount - 40) < int.Parse(sleeveline[i + 1]) && int.Parse(sleeveline[i + 1]) <= (maxShotCount - 20))
					{
						label[j].BackColor = Color.Yellow;
					}
					else if(int.Parse(sleeveline[i + 1]) > (maxShotCount - 20))
					{
						label[j].BackColor = Color.Red;
					}
					
					if(int.Parse(goodLabel[j].Text) > 0)
					{
						goodLabel[j].ForeColor = Color.Lime;
					}
					else
					{
						goodLabel[j].ForeColor = Color.Black;
					}

					if(int.Parse(furyouLabel[j].Text) > 0)
					{
						furyouLabel[j].ForeColor = Color.Red;
					}
					else
					{
						furyouLabel[j].ForeColor = Color.Black;
					}

					j++;
				}
			}

			double oknum = (double)okCount;

			int seikei_total = 0;
			if(sleeveInfo != "")
			{
				for(int i = 0; i < seikeiLabel.Length; i++)
				{
					if(seikeiLabel[i].Text != "")
					{
						seikei_total += int.Parse(seikeiLabel[i].Text);
					}
				}
			}
			double total = (double)seikei_total;
            double budomari = (double)(oknum / total * 100);

			if(total > 0)
			{
				budomari = (double)(oknum / total * 100);
			}
			else//NaN
			{
				budomari = 0;
			}

            label20.Text = budomari.ToString("F1") + "％";
			label208.Text = label22.Text = total.ToString();
			
			if(!Double.IsNaN(kadou))
			{
				kadou *= 100;
				label54.Text += kadou.ToString("F1") + "％";
			}

			label55.Text = "上限 : " + upper;
			label56.Text = "下限 : " + lower;
			label57.Text = "成型機 : " + gouki;
			label58.Text = "製品 : " + seihin;
			label207.Text = "作業者 : " + operatorName + "さん";

			if(totalSleeve > 0)
			{
                string[] nikuatsuNGline = nikuatsuNG.Split(',');
                string[] kizuNGline = kizuNG.Split(',');
                string[] butsuNGline = butsuNG.Split(',');
                string[] yakeNGline = yakeNG.Split(',');
                string[] hibicrackNGline = hibicrackNG.Split(',');
                string[] gasukizuNGline = gasukizuNG.Split(',');
                string[] houshakizuNGline = houshakizuNG.Split(',');
                string[] giratsukikumoriNGline = giratsukikumoriNG.Split(',');
                string[] hennikumendareNGline = hennikumendareNG.Split(',');
                string[] yogoreNGline = yogoreNG.Split(',');
                string[] hokoriNGline = hokoriNG.Split(',');
                string[] keijoseidoNGline = keijoseidoNG.Split(',');
                string[] etcNGline = etcNG.Split(',');
                string[] tachiageNGline = tachiageNG.Split(',');

				for(int i = 0; i < totalSleeve; i++)
				{
					nikufuLabel[i].Text = nikuatsuNGline[i];
					kizuLabel[i].Text = kizuNGline[i];
					butsuLabel[i].Text = butsuNGline[i];
					yakeLabel[i].Text = yakeNGline[i];
					hibicrackLabel[i].Text = hibicrackNGline[i];
					gasukizuLabel[i].Text = gasukizuNGline[i];
					houshsaLabel[i].Text = houshakizuNGline[i];
					giratsukiLabel[i].Text = giratsukikumoriNGline[i];
					hennikuLabel[i].Text = hennikumendareNGline[i];
					yogoreLabel[i].Text = yogoreNGline[i];
					hokoriLabel[i].Text = hokoriNGline[i];
					keijouseidoLabel[i].Text = keijoseidoNGline[i];
					sonotaLabel[i].Text = etcNGline[i];
					tachiageLabel[i].Text = tachiageNGline[i];
				}
			}
		}

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
			timer1.Enabled = false;
			timer2.Enabled = false;
			timer2.Enabled = true;//restart

		    DateTime start = e.Start;
		    DateTime end = e.End;
			SetSeikeiInfo(start, end);
        }

        private void FormCalender_Load(object sender, EventArgs e)
        {
        }

        private void FormCalender_Shown(object sender, EventArgs e)
        {
            if(formType == 0 || formType == 1)//LS,NQD
			{
				this.Width = Form1.SETDATA.shukeiWidth;
				this.Height = Form1.SETDATA.shukeiHeight;
	        }
			else if(formType == 2)//HS
			{
				this.Width = Form3.SETDATA.shukeiWidth;
				this.Height = Form3.SETDATA.shukeiHeight;
			}
			else if(formType == 3 || formType == 4)//マルチLS,マルチNQD
			{
				this.Width = Form4.SETDATA.shukeiWidth;
				this.Height = Form4.SETDATA.shukeiHeight;
			}
			else if(formType == 5)//マルチHS
			{
				this.Width = Form6.SETDATA.shukeiWidth;
				this.Height = Form6.SETDATA.shukeiHeight;
			}

			DateTime dt = DateTime.Now;
			SetSeikeiInfo(dt, dt);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
			DateTime dt = DateTime.Now;
			SetSeikeiInfo(dt, dt);
        }

        private void button2_Click(object sender, EventArgs e)
        {
			if(!shimeFlg)
			{
				shimeFlg = true;
				button2.Text = "戻　る";
				label1.Text = "範　囲";
				button1.Visible = true;
			}
			else
			{
				shimeFlg = false;
				button2.Text = "締める";
				label1.Text = "監視中";
				button1.Visible = false;
			}
        }

		private void GetEditString(string beforeStr, ref string afterStr)
		{
			string NonGouki = beforeStr.Substring(0, (beforeStr.Length - 2));
			string HeadStr = beforeStr.Substring(0, 2);
			string num = "";
			if(NonGouki.IndexOf("LS") != -1 || NonGouki.IndexOf("HS") != -1)
			{
				int len = NonGouki.Length;
				if(len == 3)
				{
					num = NonGouki.Substring(2, 1);
				}
				else if(len == 4)
				{
					num = NonGouki.Substring(2, 2);
				}
				string numdata = String.Format("{0:D2}", int.Parse(num));
				afterStr = HeadStr + " " + numdata;
			}

			if(NonGouki.IndexOf("NQD") != -1)
			{
				int len = NonGouki.Length;
				if(len == 4)
				{
					num = NonGouki.Substring(3, 1);
				}
				else if(len == 5)
				{
					num = NonGouki.Substring(3, 2);
				}
				string numdata = String.Format("{0:D2}", int.Parse(num));
				afterStr = HeadStr + " " + numdata;
			}
		}


        private void button1_Click(object sender, EventArgs e)
        {
			if(label22.Text == "0")
			{
                MessageBox.Show("データがありません", "確認");
				return;
			}

			DateTime sta_datehani = DateTime.Now;
			DateTime end_datehani = DateTime.Now;
//            DateTime sta_datehani = monthCalendar1.SelectionStart;
//            DateTime end_datehani = monthCalendar1.SelectionEnd;
//			string dailysign = sta_datehani.ToString("HHmmss");

			string edit_gouki = "";
            if(formType == 0 || formType == 1)//LS,NQD
			{
				GetEditString(gouki, ref edit_gouki);
                string symmaryStr = sta_datehani.ToString("yyyyMMdd") + "," + edit_gouki + "," + seihin + "," + label20.Text + "," +
                                    label22.Text + "," + label18.Text + "," + label19.Text + ",";

				symmaryStr += label36.Text + "," + label37.Text + "," + label38.Text + "," + label39.Text + "," + label40.Text + "," + 
								label41.Text + "," + label42.Text + "," + label43.Text + "," + label44.Text + "," + label45.Text + "," + 
								label46.Text + "," + label47.Text + "," + label48.Text;

				symmaryStr += "," + operatorName;
				symmaryStr += "," + label50.Text + "," + label51.Text + "," + label52.Text + "," + label53.Text;
				
				bool retValue = form1.WriteDailySummaryToCsv(sta_datehani, end_datehani, symmaryStr);
                if(!retValue)
                {
                    //締めの開始または終了が指定されていない。
                    label49.Text = "締め開始/終了が未指定";
                    return;
                }

                MessageBox.Show("該当日のCSVを保存しました", "終了");
	        }
			else if(formType == 2)//HS
			{
				GetEditString(gouki, ref edit_gouki);
                string symmaryStr = sta_datehani.ToString("yyyyMMdd") + "," + edit_gouki + "," + seihin + "," + label20.Text + "," +
                                    label22.Text + "," + label18.Text + "," + label19.Text + ",";

				symmaryStr += label36.Text + "," + label37.Text + "," + label38.Text + "," + label39.Text + "," + label40.Text + "," + 
								label41.Text + "," + label42.Text + "," + label43.Text + "," + label44.Text + "," + label45.Text + "," + 
								label46.Text + "," + label47.Text + "," + label48.Text;

				symmaryStr += "," + operatorName;
				symmaryStr += "," + label50.Text + "," + label51.Text + "," + label52.Text + "," + label53.Text;
				
	            bool retValue = form3.WriteDailySummaryToCsv(sta_datehani, end_datehani, symmaryStr);
                if(!retValue)
                {
                    //締めの開始または終了が指定されていない。
                    label49.Text = "締め開始/終了が未指定";
                    return;
                }

                MessageBox.Show("該当日のCSVを保存しました", "終了");
			}
			else if(formType == 3 || formType == 4)//マルチLS,マルチNQD
			{
				GetEditString(gouki, ref edit_gouki);
                string symmaryStr = sta_datehani.ToString("yyyyMMdd") + "," + edit_gouki + "," + seihin + "," + label20.Text + "," +
                                    label22.Text + "," + label18.Text + "," + label19.Text + ",";

				symmaryStr += label36.Text + "," + label37.Text + "," + label38.Text + "," + label39.Text + "," + label40.Text + "," + 
								label41.Text + "," + label42.Text + "," + label43.Text + "," + label44.Text + "," + label45.Text + "," + 
								label46.Text + "," + label47.Text + "," + label48.Text;

				symmaryStr += "," + operatorName;
				symmaryStr += "," + label50.Text + "," + label51.Text + "," + label52.Text + "," + label53.Text;
				
	            bool retValue = form4.WriteDailySummaryToCsv(sta_datehani, end_datehani, symmaryStr);
                if(!retValue)
                {
                    //締めの開始または終了が指定されていない。
                    label49.Text = "締め開始/終了が未指定";
                    return;
                }

                MessageBox.Show("該当日のCSVを保存しました", "終了");
			}
			else if(formType == 5)//マルチHS
			{
				GetEditString(gouki, ref edit_gouki);
                string symmaryStr = sta_datehani.ToString("yyyyMMdd") + "," + edit_gouki + "," + seihin + "," + label20.Text + "," +
                                    label22.Text + "," + label18.Text + "," + label19.Text + ",";

				symmaryStr += label36.Text + "," + label37.Text + "," + label38.Text + "," + label39.Text + "," + label40.Text + "," + 
								label41.Text + "," + label42.Text + "," + label43.Text + "," + label44.Text + "," + label45.Text + "," + 
								label46.Text + "," + label47.Text + "," + label48.Text;

				symmaryStr += "," + operatorName;
				symmaryStr += "," + label50.Text + "," + label51.Text + "," + label52.Text + "," + label53.Text;
				
	            bool retValue = form6.WriteDailySummaryToCsv(sta_datehani, end_datehani, symmaryStr);
                if(!retValue)
                {
                    //締めの開始または終了が指定されていない。
                    label49.Text = "締め開始/終了が未指定";
                    return;
                }

                MessageBox.Show("該当日のCSVを保存しました", "終了");
			}
        }

        private void FormCalender_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(formType == 0 || formType == 1)//LS,NQD
			{
//				Form1.SETDATA.shukeiWidth = this.Width;
//				Form1.SETDATA.shukeiHeight = this.Height;
	        }
			else if(formType == 2)//HS
			{
//				Form3.SETDATA.shukeiWidth = this.Width;
//				Form3.SETDATA.shukeiHeight = this.Height;
			}
			else if(formType == 3 || formType == 4)//マルチLS,マルチNQD
			{
//				Form4.SETDATA.shukeiWidth = this.Width;
//				Form4.SETDATA.shukeiHeight = this.Height;
			}
			else if(formType == 5)//マルチHS
			{
//				Form6.SETDATA.shukeiWidth = this.Width;
//				Form6.SETDATA.shukeiHeight = this.Height;
			}
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
			timer2.Enabled = false;
			timer1.Enabled = true;
        }

        private void alive_timer_Tick(object sender, EventArgs e)
        {
            //画面暗転阻止
            SetThreadExecutionState(ExecutionState.DisplayRequired);

            // ドラッグ操作の準備 (struct 配列の宣言)
            INPUT[] input = new INPUT[1];  // イベントを格納

            // ドラッグ操作の準備 (イベントの定義 = 相対座標へ移動)
            input[0].mi.dx = 0;  // 相対座標で0　つまり動かさない
            input[0].mi.dy = 0;  // 相対座標で0 つまり動かさない
            input[0].mi.dwFlags = MOUSEEVENTF_MOVED;

            // ドラッグ操作の実行 (イベントの生成)
            SendInput(1, input, Marshal.SizeOf(input[0]));
        }
    }
}
