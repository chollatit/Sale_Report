using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sale_Report.Model;

namespace Sale_Report.View.Master
{
    public partial class OEMPartGroupMS : Form
    {
        OEMPartMS oemPartms = new OEMPartMS();

        string processCls = "";
        public OEMPartGroupMS()
        {
            InitializeComponent();

            txtDateTime.Text = DateTime.Now.ToString();
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            txtDateTime.Text = DateTime.Now.ToString();
        }

        private void rdbAddShipto_CheckedChanged(object sender, EventArgs e)
        {
            processCls = "add";
            activeTbxItem(true);
            tbxItemCD.Focus();
            btnUpdate.Enabled = true;
        }

        private void rdbDelShipto_CheckedChanged(object sender, EventArgs e)
        {
            processCls = "del";
            btnUpdate.Enabled = true;
        }

        private void activeTbxItem(bool flg)
        {
            tbxItemCD.Enabled = flg;
        }

        private void tbxItemCD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (tbxItemCD.Text.Length > 0)
                {
                    //Find item in db
                    string itemDesc = findItemDesc(tbxItemCD.Text.Trim());
                    if (itemDesc == "")
                    {
                        MessageBox.Show("No data.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else
                    {
                        tbxItemDesc.Text = itemDesc;
                        btnUpdate.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("Please input item cd.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
        }

        private string findItemDesc(string itemCD)
        {
            try
            {
                string itemDesc = "";
                DataSet dsTmp = oemPartms.selectItemDetail(itemCD);

                if (dsTmp != null)
                {
                    if (dsTmp.Tables[0].Rows.Count > 0)
                    {
                        itemDesc = dsTmp.Tables[0].Rows[0]["i_item_desc"].ToString();
                    }
                }

                return itemDesc;
            }
            catch
            {
                return "";
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string itemCD = tbxItemCD.Text;
            string itemDesc = tbxItemDesc.Text;


        }
    }
}
