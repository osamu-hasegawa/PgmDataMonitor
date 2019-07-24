using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PgmDataMonitor
{
    public partial class Form7 : Form
    {
		public string EditInfo = "";
		public string CauseInfo = "";
		public string SeikeiInfo = "";
		double upper = 0;
		double lower = 0;
		double initData = 0;

        public Form7()
        {
            InitializeComponent();
        }

		public void SetInfo(string sleeve, string shotCount, string nikuUplimit, string nikuDnlimit, string nikuEdit, string resultCause)
		{
			comboBox1.Items.Add("なし");
			comboBox2.Items.Add("なし");
			comboBox3.Items.Add("なし");
			comboBox4.Items.Add("なし");
			comboBox5.Items.Add("なし");
			comboBox6.Items.Add("なし");
			for(int i = 0; i < Form1.SETDATA.ngCause.Length; i++)//不良原因
			{
				comboBox1.Items.Add(Form1.SETDATA.ngCause[i]);
				comboBox2.Items.Add(Form1.SETDATA.ngCause[i]);
				comboBox3.Items.Add(Form1.SETDATA.ngCause[i]);
				comboBox4.Items.Add(Form1.SETDATA.ngCause[i]);
				comboBox5.Items.Add(Form1.SETDATA.ngCause[i]);
				comboBox6.Items.Add(Form1.SETDATA.ngCause[i]);
			}

			this.Text = "マルチCAV　肉厚入力して下さい";
			label3.Text = sleeve;
			label1.Text = nikuUplimit;
			label2.Text = nikuDnlimit;
			numericUpDown7.Text = shotCount;

			upper = double.Parse(nikuUplimit);
			lower = double.Parse(nikuDnlimit);
			initData = (upper + lower) / 2;

			NumericUpDown [] numeric = new NumericUpDown[] {numericUpDown1, numericUpDown2, numericUpDown3, numericUpDown4, numericUpDown5, numericUpDown6};
			ComboBox [] combo = new ComboBox[] {comboBox1, comboBox2, comboBox3, comboBox4, comboBox5, comboBox6};
			CheckBox [] check = new CheckBox[] {checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6};
			GroupBox [] group = new GroupBox[] {groupBox3, groupBox5, groupBox7, groupBox9, groupBox11, groupBox13};

			if(nikuEdit == ",,,,,")
			{
				for(int i = 0; i < numeric.Length; i++)
				{
					numeric[i].Enabled = false;
					numeric[i].Value = (decimal)0;
				}

				for(int i = 0; i < combo.Length; i++)
				{
					combo[i].Enabled = false;
					combo[i].Text = "なし";
				}

				for(int i = 0; i < check.Length; i++)
				{
					check[i].Checked = false;
				}

				for(int i = 0; i < group.Length; i++)
				{
					group[i].BackColor = Color.White;
				}

			}
			else
			{
				string[] strline = nikuEdit.Split(',');
				string[] causeline = resultCause.Split(',');

				for(int i = 0; i < numeric.Length; i++)
				{
					if(strline[i] == "-" || strline[i] == "")
					{
						numeric[i].Text = "0.000";
					}
					else
					{
						numeric[i].Text = strline[i];
					}

					if(causeline[i] == "-" || causeline[i] == "OK" || causeline[i] == "")
					{
						causeline[i] = "なし";
					}
					combo[i].Text = causeline[i];

					if(numeric[i].Text == "0.000")
					{
						check[i].Checked = false;
						numeric[i].Enabled = false;
						combo[i].Enabled = false;

						group[i].BackColor = Color.White;
					}
					else
					{
						check[i].Checked = true;
						numeric[i].Enabled = true;
						combo[i].Enabled = true;

						if(causeline[i] == "なし")
						{
							group[i].BackColor = Color.Lime;
						}
						else
						{
							group[i].BackColor = Color.Red;
						}
					}
				}
			}
		}

        private void button1_Click(object sender, EventArgs e)
        {

			string nikuatsu = "";
            string cause = "";
            string [] niku_data = new string [6];
            string [] cause_data = new string [6];

			string sleeve = string.Format("スリーブ　：　{0}", label3.Text);
			sleeve += "\r\n";

			string shotCount = string.Format("ショット数　：　{0}", numericUpDown7.Text);
            shotCount += "\r\n";

			if(checkBox1.Checked)
			{
				nikuatsu = string.Format("Cav1　：　{0}", numericUpDown1.Text);
				cause = string.Format("Cav1不良原因　：　{0}", comboBox1.Text);
				niku_data[0] = numericUpDown1.Text;
				cause_data[0] = comboBox1.Text;
			}
			else
			{
				nikuatsu = string.Format("Cav1　：　-");
				cause = string.Format("Cav1不良原因　：　-");
				niku_data[0] = "-";
				cause_data[0] = "-";
			}
			nikuatsu += "\r\n";
			cause += "\r\n";

			if(checkBox2.Checked)
			{
				nikuatsu += string.Format("Cav2　：　{0}", numericUpDown2.Text);
				cause += string.Format("Cav2不良原因　：　{0}", comboBox2.Text);
				niku_data[1] = numericUpDown2.Text;
				cause_data[1] = comboBox2.Text;
			}
			else
			{
				nikuatsu += string.Format("Cav2　：　-");
				cause += string.Format("Cav2不良原因　：　-");
				niku_data[1] = "-";
				cause_data[1] = "-";
			}
			nikuatsu += "\r\n";
			cause += "\r\n";

			if(checkBox3.Checked)
			{
				nikuatsu += string.Format("Cav3　：　{0}", numericUpDown3.Text);
				cause += string.Format("Cav3不良原因　：　{0}", comboBox3.Text);
				niku_data[2] = numericUpDown3.Text;
				cause_data[2] = comboBox3.Text;
			}
			else
			{
				nikuatsu += string.Format("Cav3　：　-");
				cause += string.Format("Cav3不良原因　：　-");
				niku_data[2] = "-";
				cause_data[2] = "-";
			}
			nikuatsu += "\r\n";
			cause += "\r\n";

			if(checkBox4.Checked)
			{
				nikuatsu += string.Format("Cav4　：　{0}", numericUpDown4.Text);
				cause += string.Format("Cav4不良原因　：　{0}", comboBox4.Text);
				niku_data[3] = numericUpDown4.Text;
				cause_data[3] = comboBox4.Text;
			}
			else
			{
				nikuatsu += string.Format("Cav4　：　-");
				cause += string.Format("Cav4不良原因　：　-");
				niku_data[3] = "-";
				cause_data[3] = "-";
			}
			nikuatsu += "\r\n";
			cause += "\r\n";

			if(checkBox5.Checked)
			{
				nikuatsu += string.Format("Cav5　：　{0}", numericUpDown5.Text);
				cause += string.Format("Cav5不良原因　：　{0}", comboBox5.Text);
				niku_data[4] = numericUpDown5.Text;
				cause_data[4] = comboBox5.Text;
			}
			else
			{
				nikuatsu += string.Format("Cav5　：　-");
				cause += string.Format("Cav5不良原因　：　-");
				niku_data[4] = "-";
				cause_data[4] = "-";
			}
			nikuatsu += "\r\n";
			cause += "\r\n";

			if(checkBox6.Checked)
			{
				nikuatsu += string.Format("Cav6　：　{0}", numericUpDown6.Text);
				cause += string.Format("Cav6不良原因　：　{0}", comboBox6.Text);
				niku_data[5] = numericUpDown6.Text;
				cause_data[5] = comboBox6.Text;
			}
			else
			{
				nikuatsu += string.Format("Cav6　：　-");
				cause += string.Format("Cav6不良原因　：　-");
				niku_data[5] = "-";
				cause_data[5] = "-";
			}
			nikuatsu += "\r\n";
			cause += "\r\n";

			cause += "\r\n";

			string combine = sleeve + shotCount + nikuatsu + cause + "で間違いありませんか？";

			DialogResult result = MessageBox.Show(combine, "肉厚測定　入力・編集", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{   
				EditInfo = niku_data[0] + "," + niku_data[1] + "," + niku_data[2] + "," + niku_data[3] + "," + niku_data[4] + "," + niku_data[5];
				CauseInfo = cause_data[0] + "," + cause_data[1] + "," + cause_data[2] + "," + cause_data[3] + "," + cause_data[4] + "," + cause_data[5];
				SeikeiInfo = numericUpDown7.Text + ",";
                this.Close();
	        }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
			if(!checkBox1.Checked)
			{
				return;
			}

			double nikuData = (double)numericUpDown1.Value;
			if(nikuData < lower || upper < nikuData)
			{
				comboBox1.Text = "肉厚不良";
			}
			else
			{
				comboBox1.Text = "なし";
			}
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
			if(!checkBox2.Checked)
			{
				return;
			}

			double nikuData = (double)numericUpDown2.Value;
			if(nikuData < lower || upper < nikuData)
			{
				comboBox2.Text = "肉厚不良";
			}
			else
			{
				comboBox2.Text = "なし";
			}
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
			if(!checkBox3.Checked)
			{
				return;
			}

			double nikuData = (double)numericUpDown3.Value;
			if(nikuData < lower || upper < nikuData)
			{
				comboBox3.Text = "肉厚不良";
			}
			else
			{
				comboBox3.Text = "なし";
			}
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
			if(!checkBox4.Checked)
			{
				return;
			}

			double nikuData = (double)numericUpDown4.Value;
			if(nikuData < lower || upper < nikuData)
			{
				comboBox4.Text = "肉厚不良";
			}
			else
			{
				comboBox4.Text = "なし";
			}
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
			if(!checkBox5.Checked)
			{
				return;
			}

			double nikuData = (double)numericUpDown5.Value;
			if(nikuData < lower || upper < nikuData)
			{
				comboBox5.Text = "肉厚不良";
			}
			else
			{
				comboBox5.Text = "なし";
			}
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
			if(!checkBox6.Checked)
			{
				return;
			}

			double nikuData = (double)numericUpDown6.Value;
			if(nikuData < lower || upper < nikuData)
			{
				comboBox6.Text = "肉厚不良";
			}
			else
			{
				comboBox6.Text = "なし";
			}
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox1.Checked)
			{
				comboBox1.Enabled = true;
				numericUpDown1.Enabled = true;

				if(numericUpDown1.Text != "0.000")
				{
					return;
				}

				comboBox1.Text = "なし";
				if(label1.Text != "" && label2.Text != "")
				{
					numericUpDown1.DecimalPlaces = 3;
					numericUpDown1.Value = (decimal)initData;
				}
				else
				{
					numericUpDown1.Value = (decimal)0;
				}
                groupBox3.BackColor = Color.Lime;
			}
			else
			{
				comboBox1.Enabled = false;
				comboBox1.Text = "なし";
				numericUpDown1.Enabled = false;
				numericUpDown1.Value = (decimal)0;
                groupBox3.BackColor = Color.White;
			}
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox2.Checked)
			{
				comboBox2.Enabled = true;
				numericUpDown2.Enabled = true;

				if(numericUpDown2.Text != "0.000")
				{
					return;
				}

				comboBox2.Text = "なし";
				if(label1.Text != "" && label2.Text != "")
				{
					numericUpDown2.DecimalPlaces = 3;
					numericUpDown2.Value = (decimal)initData;
				}
				else
				{
					numericUpDown2.Value = (decimal)0;
				}
                groupBox5.BackColor = Color.Lime;
			}
			else
			{
				comboBox2.Enabled = false;
				comboBox2.Text = "なし";
				numericUpDown2.Enabled = false;
				numericUpDown2.Value = (decimal)0;
                groupBox5.BackColor = Color.White;
			}
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox3.Checked)
			{
				comboBox3.Enabled = true;
				numericUpDown3.Enabled = true;

				if(numericUpDown3.Text != "0.000")
				{
					return;
				}

				comboBox3.Text = "なし";
				if(label1.Text != "" && label2.Text != "")
				{
					numericUpDown3.DecimalPlaces = 3;
					numericUpDown3.Value = (decimal)initData;
				}
				else
				{
					numericUpDown3.Value = (decimal)0;
				}
                groupBox7.BackColor = Color.Lime;
			}
			else
			{
				comboBox3.Enabled = false;
				comboBox3.Text = "なし";
				numericUpDown3.Enabled = false;
				numericUpDown3.Value = (decimal)0;
                groupBox7.BackColor = Color.White;
			}
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox4.Checked)
			{
				comboBox4.Enabled = true;
				numericUpDown4.Enabled = true;

				if(numericUpDown4.Text != "0.000")
				{
					return;
				}

				comboBox4.Text = "なし";
				if(label1.Text != "" && label2.Text != "")
				{
					numericUpDown4.DecimalPlaces = 3;
					numericUpDown4.Value = (decimal)initData;
				}
				else
				{
					numericUpDown4.Value = (decimal)0;
				}
                groupBox9.BackColor = Color.Lime;
			}
			else
			{
				comboBox4.Enabled = false;
				comboBox4.Text = "なし";
				numericUpDown4.Enabled = false;
				numericUpDown4.Value = (decimal)0;
                groupBox9.BackColor = Color.White;
			}
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox5.Checked)
			{
				comboBox5.Enabled = true;
				numericUpDown5.Enabled = true;

				if(numericUpDown5.Text != "0.000")
				{
					return;
				}

				comboBox5.Text = "なし";
				if(label1.Text != "" && label2.Text != "")
				{
					numericUpDown5.DecimalPlaces = 3;
					numericUpDown5.Value = (decimal)initData;
				}
				else
				{
					numericUpDown5.Value = (decimal)0;
				}
                groupBox11.BackColor = Color.Lime;
			}
			else
			{
				comboBox5.Enabled = false;
				comboBox5.Text = "なし";
				numericUpDown5.Enabled = false;
				numericUpDown5.Value = (decimal)0;
                groupBox11.BackColor = Color.White;
			}
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
			if(checkBox6.Checked)
			{
				comboBox6.Enabled = true;
				numericUpDown6.Enabled = true;

				if(numericUpDown6.Text != "0.000")
				{
					return;
				}

				comboBox6.Text = "なし";
				if(label1.Text != "" && label2.Text != "")
				{
					numericUpDown6.DecimalPlaces = 3;
					numericUpDown6.Value = (decimal)initData;
				}
				else
				{
					numericUpDown6.Value = (decimal)0;
				}
                groupBox13.BackColor = Color.Lime;
			}
			else
			{
				comboBox6.Enabled = false;
				comboBox6.Text = "なし";
				numericUpDown6.Enabled = false;
				numericUpDown6.Value = (decimal)0;
                groupBox13.BackColor = Color.White;
			}
        }

        private void button2_Click(object sender, EventArgs e)
        {
			string nikuUpLimit = label1.Text;
			string nikuLoLimit = label2.Text;
			string nikuEdit = "";
			string resultCause = "なし";

            //登録画面の表示
            string justSleeve = label3.Text;
			string currentShotCount = numericUpDown7.Text;
            Form5 form5 = new Form5();
			form5.SetInfo(justSleeve, currentShotCount, nikuUpLimit, nikuLoLimit, nikuEdit, resultCause);
            form5.ShowDialog();

            string InfoStr = form5.EditInfo;
            if(InfoStr == "")
            {
				return;
			}

			string[] strline = InfoStr.Split(',');
			string nikuAtai = strline[0];
			string ngCause = strline[1];
			numericUpDown7.Text = strline[2];

			NumericUpDown [] numeric = new NumericUpDown[] {numericUpDown1, numericUpDown2, numericUpDown3, numericUpDown4, numericUpDown5, numericUpDown6};
			ComboBox [] combo = new ComboBox[] {comboBox1, comboBox2, comboBox3, comboBox4, comboBox5, comboBox6};
			CheckBox [] check = new CheckBox[] {checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6};

			for(int i = 0; i < numeric.Length; i++)
			{
				check[i].Checked = true;
				numeric[i].Text = nikuAtai;
				combo[i].Text = ngCause;
			}
			
			MessageBox.Show("他のCavにも反映しました。保存ボタンを押して下さい", "まとめて入力", MessageBoxButtons.OK);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.Text == "" || comboBox1.Text == "なし")
            {
                groupBox3.BackColor = Color.Lime;
            }
            else
            {
                groupBox3.BackColor = Color.Red;
			}
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox2.Text == "" || comboBox2.Text == "なし")
            {
                groupBox5.BackColor = Color.Lime;
            }
            else
            {
                groupBox5.BackColor = Color.Red;
			}
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox3.Text == "" || comboBox3.Text == "なし")
            {
                groupBox7.BackColor = Color.Lime;
            }
            else
            {
                groupBox7.BackColor = Color.Red;
			}
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox4.Text == "" || comboBox4.Text == "なし")
            {
                groupBox9.BackColor = Color.Lime;
            }
            else
            {
                groupBox9.BackColor = Color.Red;
			}
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox5.Text == "" || comboBox5.Text == "なし")
            {
                groupBox11.BackColor = Color.Lime;
            }
            else
            {
                groupBox11.BackColor = Color.Red;
			}
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox6.Text == "" || comboBox6.Text == "なし")
            {
                groupBox13.BackColor = Color.Lime;
            }
            else
            {
                groupBox13.BackColor = Color.Red;
			}
        }
    }
}
