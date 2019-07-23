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

        DataSet dsOemPartList = null;
        string processCls = "";

        public OEMPartGroupMS()
        {
            InitializeComponent();

            initDGVListOemPart();
            txtDateTime.Text = DateTime.Now.ToString();
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            txtDateTime.Text = DateTime.Now.ToString();
        }

        private void initDGVListOemPart()
        {
            dsOemPartList = new DataSet();
            dsOemPartList = oemPartms.selectOemPartList();

            if (dsOemPartList == null)
            {
                MessageBox.Show("No data.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                dgvListOEMItemCD.DataSource = dsOemPartList.Tables[0];
            }
        }

        private void clearTbxItem()
        {
            tbxItemCD.Text = "";
            tbxItemDesc.Text = "";
        }

        private void rdbAddShipto_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbAddShipto.Checked)
            {
                processCls = "add";
                activeTbxItem(true);
                tbxItemCD.Focus();
                btnUpdate.Enabled = true;
            }
        }

        private void rdbDelShipto_CheckedChanged(object sender, EventArgs e)
        {
            processCls = "del";
            activeTbxItem(false);
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

            if (processCls == "add")
            {
                string strCmd = "trim(i_item_cd) = '" + itemCD + "'";
                DataRow[] dr = dsOemPartList.Tables[0].Select(strCmd);

                if (dr.Count() > 0)
                {
                    MessageBox.Show("This item cd already in list below.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    if (tbxItemDesc.Text.Count() > 0)
                    {
                        int result = oemPartms.insertOEMItem(itemCD, itemDesc);

                        if (result > 0)
                        {
                            MessageBox.Show("Add complete.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            clearTbxItem();
                            tbxItemCD.Focus();
                            initDGVListOemPart();
                        }
                        else
                        {
                            MessageBox.Show("Add incomplete.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please input valid item cd.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
            }
            else if (processCls == "del")
            {
                DialogResult dialogResult = MessageBox.Show("Do you want to delete this item cd?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {
                    int result = oemPartms.deleteItemCD(itemCD);

                    if (result > 0)
                    {
                        MessageBox.Show("Delete complete.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        clearTbxItem();
                        initDGVListOemPart();
                    }
                    else
                    {
                        MessageBox.Show("Delete incomplete.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (tbxItemCD.Text.Count() > 0)
            {
                clearTbxItem();
                tbxItemCD.Enabled = false;
            }
            else if (processCls != "")
            {
                rdbAddShipto.Checked = false;
                rdbDelShipto.Checked = false;
                tbxItemCD.Enabled = false;
                processCls = "";
            }
        }

        private void dgvListOEMItemCD_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (processCls == "del")
            {
                string itemCD = dgvListOEMItemCD.Rows[e.RowIndex].Cells["i_item_cd"].Value.ToString();
                string itemDesc = dgvListOEMItemCD.Rows[e.RowIndex].Cells["i_item_desc"].Value.ToString();

                tbxItemCD.Text = itemCD;
                tbxItemDesc.Text = itemDesc;
            }
        }
    }
}
