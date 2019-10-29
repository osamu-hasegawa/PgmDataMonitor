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
    public partial class Form5 : Form
    {
		public string EditInfo = "";
		double upper = 0;
		double lower = 0;
        public Form5()
        {
            InitializeComponent();
			this.MaximizeBox = false;
			alive_timer.Enabled = true;
        }

		public void SetInfo(string sleeve, string shotCount, string nikuUplimit, string nikuDnlimit, string nikuEdit, string resultCause)
		{
			this.Text = "編集画面";
			label5.Text = sleeve;
			label1.Text = nikuUplimit;
			label2.Text = nikuDnlimit;
			numericUpDown2.Text = shotCount;

			upper = double.Parse(nikuUplimit);
			lower = double.Parse(nikuDnlimit);
			double data = (upper + lower) / 2;

			if(nikuEdit == "")
			{
				if(nikuUplimit != "" && nikuDnlimit != "")
				{
					numericUpDown1.DecimalPlaces = 3;
					numericUpDown1.Value = (decimal)data;
				}
				else
				{
					numericUpDown1.Value = (decimal)0;
				}
			}
			else
			{
				numericUpDown1.Text = nikuEdit;
			}
			
			comboBox1.Items.Add("なし");
			for(int i = 0; i < Form1.SETDATA.ngCause.Length; i++)//不良原因
			{
				comboBox1.Items.Add(Form1.SETDATA.ngCause[i]);
			}
			
			if(resultCause == "-" || resultCause == "OK" || resultCause == "なし")
			{
				comboBox1.Text = "なし";
                groupBox4.BackColor = Color.Lime;
			}
			else
			{
				comboBox1.Text = resultCause;//NGの時
                groupBox4.BackColor = Color.Red;
			}
			
		}

        private void button1_Click(object sender, EventArgs e)
        {
			string sleeve = string.Format("スリーブ　：　{0}", label5.Text);
            sleeve += "\r\n";

			string shotCount = string.Format("ショット数　：　{0}", numericUpDown2.Text);
            shotCount += "\r\n";
			
			string nikuatsu = string.Format("肉厚　：　{0}", numericUpDown1.Text);
            nikuatsu += "\r\n";

            string cause = string.Format("不良原因　：　{0}", comboBox1.Text);
            cause += "\r\n";

            string combine = sleeve + shotCount + nikuatsu + cause + "で間違いありませんか？";

            DialogResult result = MessageBox.Show(combine, "肉厚測定　入力・編集", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{   
				EditInfo = numericUpDown1.Text + "," + comboBox1.Text + "," + numericUpDown2.Text;
                this.Close();
	        }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
			//タイマーリスタート
			alive_timer.Enabled = false;
			alive_timer.Enabled = true;

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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
			//タイマーリスタート
			alive_timer.Enabled = false;
			alive_timer.Enabled = true;

            if(comboBox1.Text == "" || comboBox1.Text == "なし")
            {
                groupBox4.BackColor = Color.Lime;
            }
            else
            {
                groupBox4.BackColor = Color.Red;
			}
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
			//タイマーリスタート
			alive_timer.Enabled = false;
			alive_timer.Enabled = true;
        }

        private void alive_timer_Tick(object sender, EventArgs e)
        {
			this.Close();
        }
    }
}
