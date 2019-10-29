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
    public partial class FormSeikeiList : Form
    {
		FormConfirmWindow formConfirmWindow = null;
		string[] stringName;

        public FormSeikeiList()
        {
            InitializeComponent();
        }

        private void FormSeikeiList_Load(object sender, EventArgs e)
        {
        }

		public void SetFileList(string file_name_comb, string file_time_comb)
		{

            ColumnHeader columnType;
			columnType = new ColumnHeader();
			columnType.Text = "型式名";

            ColumnHeader columnTime;
			columnTime = new ColumnHeader();
			columnTime.Text = "最終成型日時";

            ColumnHeader[] colHeaderRegValue = {columnType, columnTime};
            listView1.Columns.AddRange(colHeaderRegValue);

			stringName = file_name_comb.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
			string[] stringTime = file_time_comb.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

			for(int i = 0; i < stringName.Length; i++)
			{
				string type_only = "";
				if(stringName[i].IndexOf("_monthly") >= 0)//monthlyバックアップCSV
				{
					type_only = stringName[i].Substring(0, 7);
				}
				else
				{
					type_only = stringName[i].Substring(7, 7);
				}

				string[] item = {type_only, stringTime[i]};
				listView1.Items.Add(new ListViewItem(item));
			}

			this.ActiveControl = this.listView1;
            listView1.FullRowSelect = true;
//            listView1.Sorting = SortOrder.Ascending;
            listView1.HideSelection = false;

            listView1.GridLines = true;
            listView1.View = View.Details;

            listView1.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
		}

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
			    int index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置
				string tar_file = stringName[index];

				if(formConfirmWindow == null || formConfirmWindow.IsDisposed)
				{
					formConfirmWindow = new FormConfirmWindow();
                    formConfirmWindow.SetTarget(tar_file);
                    formConfirmWindow.ShowDialog(this);

	                formConfirmWindow.Dispose();
		        }
			}

        }
    }
}
