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
    public partial class Form2 : Form
    {
        private const UInt32 SC_CLOSE = 0x0000F060;
        private const UInt32 MF_BYCOMMAND = 0x00000000;

        public Form2()
        {
            InitializeComponent();
			
            Form1.SETDATA.load(ref Form1.SETDATA);//XMLを読み出してデータに格納

			if(Form1.SETDATA.selectedMachine != "")//直前に起動している
			{
				switch(Form1.SETDATA.machineType)
				{
					case 0://LS成型機
					case 1://NQD成型機
		                Form1 form1 = new Form1();
		                form1.Show();
						break;
					case 2://HS成型機
		                Form3 form3 = new Form3();
		                form3.Show();
						break;
					case 3://LSマルチ成型機
					case 4://NQDマルチ成型機
		                Form4 form4 = new Form4();
		                form4.Show();
						break;
					case 5://HSマルチ成型機
		                Form6 form6 = new Form6();
		                form6.Show();
						break;
					default:
						break;
				}
				
				timer1.Enabled = true;
			}
			else//初めて起動する
			{
	            this.MaximizeBox = false;//最大化ボタン無効
	            label1.Text = "シングルとマルチを" + "\r\n" + "間違えないように選択して下さい";

	            for(int i = 0; i < Form1.SETDATA.machineKind.Length; i++)
	            {
	                comboBox1.Items.Add(Form1.SETDATA.machineKind[i]);
	            }
	        }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string seikeiki = string.Format("「{0}」で間違いありませんか？",comboBox1.Text);
            DialogResult result = MessageBox.Show(seikeiki, "PGM成型機　選択", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{
                Form1.SETDATA.machineType = comboBox1.SelectedIndex;
                Form1.SETDATA.save(Form1.SETDATA);
				switch(comboBox1.SelectedIndex)
				{
				case 0://LS成型機
				case 1://NQD成型機
	                Form1 form1 = new Form1();
	                form1.Show();
					break;
				case 2://HS成型機
	                Form3 form3 = new Form3();
	                form3.Show();
					break;
				case 3://LSマルチ成型機
				case 4://NQDマルチ成型機
	                Form4 form4 = new Form4();
	                form4.Show();
					break;
				case 5://HSマルチ成型機
	                Form6 form6 = new Form6();
	                form6.Show();
					break;
				default:
					break;
				}

                this.Hide();
            }
		}

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			SettingForm setform = new SettingForm();
			setform.Show();
			this.Hide();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
			timer1.Enabled = false;
            this.Hide();
        }
    }
}
