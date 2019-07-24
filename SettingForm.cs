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
    public partial class SettingForm : Form
    {
        public SettingForm()
        {
            InitializeComponent();

			//XMLから読み込み
            Form1.SETDATA.load(ref Form1.SETDATA);

			for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
			{
                comboBox5.Items.Add(Form1.SETDATA.seihinName[i]);
			}
        }

        private void SettingForm_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
			string seihin = comboBox1.Text + comboBox2.Text + comboBox3.Text + numericUpDown1.Text + numericUpDown2.Text + numericUpDown3.Text + numericUpDown4.Text + comboBox4.Text;

			//型式チェック
			if(seihin == "0000" || numericUpDown5.Text == "0.000" || numericUpDown6.Text == "0.000")
			{
				MessageBox.Show("データが不正です", "製品名　登録", MessageBoxButtons.OK);
				return;
			}

			//重複チェック
			for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
			{
				if(Form1.SETDATA.seihinName[i] == seihin)
				{
					MessageBox.Show("既に登録済です。下の画面から編集して下さい", "製品名　登録", MessageBoxButtons.OK);
					return;
				}
			}

            string mes = string.Format("製品名：「{0}」", seihin);
			mes += "\r\n";

            mes += string.Format("肉厚上限　：　{0}", numericUpDown5.Text);
			mes += "\r\n";
            mes += string.Format("肉厚下限　：　{0}", numericUpDown6.Text);
			mes += "\r\n";

            mes += string.Format("加熱時間上限　：　{0}", numericUpDown9.Text);
			mes += "\r\n";
            mes += string.Format("加熱時間下限　：　{0}", numericUpDown10.Text);
			mes += "\r\n";

            mes += string.Format("で登録します。間違いありませんか？");

            DialogResult result = MessageBox.Show(mes, "製品名　登録", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{
                //追加
				for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
				{
					if(Form1.SETDATA.seihinName[i] != "")
					{
						continue;
					}
					
					Form1.SETDATA.seihinName[i] = seihin;
					Form1.SETDATA.nikuUpper[i] = numericUpDown5.Text;
					Form1.SETDATA.nikuLower[i] = numericUpDown6.Text;

					Form1.SETDATA.KaatsuJikanUpper[i] = numericUpDown9.Text;
					Form1.SETDATA.KaatsuJikanLower[i] = numericUpDown10.Text;
					break;
				}

				//XMLに書き込み
	            Form1.SETDATA.save(Form1.SETDATA);

				//再度、読込み
				comboBox5.Items.Clear();
				for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
				{
	                comboBox5.Items.Add(Form1.SETDATA.seihinName[i]);
				}

	            MessageBox.Show("新規登録しました");
			}
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
			for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)//肉厚検索、表示
			{
				if(Form1.SETDATA.seihinName[i] == comboBox5.Text)
				{
					numericUpDown7.Text = Form1.SETDATA.nikuUpper[i];
					numericUpDown8.Text = Form1.SETDATA.nikuLower[i];
					break;
				}
			}

			for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
			{
				if(Form1.SETDATA.seihinName[i] == comboBox5.Text)
				{
					numericUpDown11.Text = Form1.SETDATA.KaatsuJikanUpper[i];
					numericUpDown12.Text = Form1.SETDATA.KaatsuJikanLower[i];
				}
			}

        }

        private void button2_Click(object sender, EventArgs e)
        {
			if(comboBox5.Text == "" || numericUpDown7.Text == "0.000" || numericUpDown8.Text == "0.000")
			{
				MessageBox.Show("データが不正です", "製品名　登録", MessageBoxButtons.OK);
				return;
			}


            string mes = string.Format("製品名：「{0}」", comboBox5.Text);
			mes += "\r\n";

            mes += string.Format("肉厚上限　：　{0}", numericUpDown7.Text);
			mes += "\r\n";
            mes += string.Format("肉厚下限　：　{0}", numericUpDown8.Text);
			mes += "\r\n";

            mes += string.Format("加熱時間上限　：　{0}", numericUpDown11.Text);
			mes += "\r\n";
            mes += string.Format("加熱時間下限　：　{0}", numericUpDown12.Text);
			mes += "\r\n";

            mes += string.Format("に更新します。間違いありませんか？");

            DialogResult result = MessageBox.Show(mes, "製品名　更新", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{
                //更新
				for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
				{
					if(Form1.SETDATA.seihinName[i] == comboBox5.Text)
					{
						Form1.SETDATA.nikuUpper[i] = numericUpDown7.Text;
						Form1.SETDATA.nikuLower[i] = numericUpDown8.Text;

						Form1.SETDATA.KaatsuJikanUpper[i] = numericUpDown11.Text;
						Form1.SETDATA.KaatsuJikanLower[i] = numericUpDown12.Text;
						break;
					}
				}

				//XMLに書き込み
	            Form1.SETDATA.save(Form1.SETDATA);
	            MessageBox.Show("更新しました");
			}
        }

    }
}
