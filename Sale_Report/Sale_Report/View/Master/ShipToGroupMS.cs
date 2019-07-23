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
    public partial class ShipToGroupMS : Form
    {
        ShiptoMS shiptoms = new ShiptoMS();
        DataSet dsShiptoCD = new DataSet();
        DataSet dsShiptoMS = new DataSet();

        string procressCls = "";
        string groupCls = "";
        string strShiptoCD = "";
        string strShiptoDesc = "";

        public ShipToGroupMS()
        {
            InitializeComponent();

            txtDateTime.Text = DateTime.Now.ToString();
            timer1.Start();

            initComboboxShiptoMS();
            initListShiptoAlreadyAdded();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            txtDateTime.Text = DateTime.Now.ToString();
        }
        private void activeGroupCls(bool flg)
        {
            rdbOEMCls.Enabled = flg;
            rdbPMSPCls.Enabled = flg;
            rdbOtherCls.Enabled = flg;
        }

        private void activeShiptoCD(bool flg)
        {
            cmbShiptoMS.Enabled = flg;
        }

        private void activeUpdateBtn(bool flg)
        {
            btnUpdate.Enabled = flg;
        }

        private void rdbAddShipto_Click(object sender, EventArgs e)
        {
            procressCls = "add";
            activeGroupCls(true);
            cmbShiptoMS.Visible = true;
            tbxShiptoCD.Visible = false;
        }

        private void rdbChgShipto_Click(object sender, EventArgs e)
        {
            procressCls = "chg";
            activeGroupCls(true);
        }

        private void rdbDelShipto_Click(object sender, EventArgs e)
        {
            procressCls = "del";
            activeGroupCls(false);
            cmbShiptoMS.Visible = false;
            tbxShiptoCD.Visible = true;
            btnUpdate.Enabled = true;
        }

        private void rdbOEMCls_Click(object sender, EventArgs e)
        {
            groupCls = "oem";
            activeShiptoCD(true);
            activeUpdateBtn(true);
        }

        private void rdbPMSPCls_Click(object sender, EventArgs e)
        {
            groupCls = "pmsp";
            activeShiptoCD(true);
            activeUpdateBtn(true);
        }

        private void rdbOtherCls_Click(object sender, EventArgs e)
        {
            groupCls = "other";
            activeShiptoCD(true);
            activeUpdateBtn(true);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (cmbShiptoMS.SelectedIndex != 0)
            {
                cmbShiptoMS.SelectedIndex = 0;
                tbxShiptoDesc.Text = "";
                activeUpdateBtn(false);
                activeShiptoCD(false);
            }
            else
            {
                if (groupCls != "")
                {
                    groupCls = "";
                    rdbOEMCls.Checked = false;
                    rdbPMSPCls.Checked = false;
                    activeGroupCls(false);
                    activeUpdateBtn(false);
                    activeShiptoCD(false);
                }
                else
                {
                    if (procressCls != "")
                    {
                        procressCls = "";
                        rdbAddShipto.Checked = false;
                        rdbChgShipto.Checked = false;
                        rdbDelShipto.Checked = false;
                    }
                    else
                    {
                        this.Close();
                    }
                }
            }
        }

        private void initComboboxShiptoMS()
        {
            dsShiptoCD = shiptoms.selectShiptoCD();

            cmbShiptoMS.DisplayMember = "Text";
            cmbShiptoMS.ValueMember = "Value";

            cmbShiptoMS.Items.Add(new { Text = "Please select ship to", Value = "Please select ship to" });
            for (int i = 0; i < dsShiptoCD.Tables[0].Rows.Count; i++)
            {
                cmbShiptoMS.Items.Add(new
                {
                    Text = dsShiptoCD.Tables[0].Rows[i]["i_dl_cd"].ToString(),
                    Value = dsShiptoCD.Tables[0].Rows[i]["i_dl_arg_desc"].ToString()
                });
            }

            cmbShiptoMS.SelectedIndex = 0;
        }

        private void initListShiptoAlreadyAdded()
        {
            dsShiptoMS = shiptoms.selectShiptoMS();
            dgvListShipto.DataSource = dsShiptoMS.Tables[0];
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (procressCls == "add")
            {
                if (cmbShiptoMS.SelectedIndex != 0)
                {
                    if (!isHasInListShipto(cmbShiptoMS.Text))
                    {

                        string strGroupCD = "";
                        if (groupCls == "oem") strGroupCD = "01";
                        else if (groupCls == "pmsp") strGroupCD = "02";
                        else if (groupCls == "oem service") strGroupCD = "03";
                        else if (groupCls == "other") strGroupCD = "04";

                        int result = shiptoms.insertShiptoMS(strGroupCD, groupCls.ToUpper(), strShiptoCD, strShiptoDesc);
                        if (result > 0)
                        {
                            MessageBox.Show("Add ship to completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            initListShiptoAlreadyAdded();

                            cmbShiptoMS.SelectedIndex = 0;
                            cmbShiptoMS.Focus();
                            tbxShiptoDesc.Text = "";
                            return;
                        }
                        else
                        {
                            MessageBox.Show("Add ship to incompleted.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                    }
                    else
                    {
                        MessageBox.Show("This ship to is already in list below.\nPleas select other.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please select ship to.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else if (procressCls == "del")
            {
                if (strShiptoCD != "")
                {
                    DialogResult dialog = MessageBox.Show("Do yo want to delete this ship to?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialog == DialogResult.Yes)
                    {
                        int result = shiptoms.deleteShiptoMS(strShiptoCD);
                        if (result > 0)
                        {
                            MessageBox.Show("Delete ship to completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            initListShiptoAlreadyAdded();

                            cmbShiptoMS.SelectedIndex = 0;
                            cmbShiptoMS.Focus();
                            tbxShiptoCD.Text = "";
                            tbxShiptoDesc.Text = "";
                            return;
                        }
                        else
                        {
                            MessageBox.Show("Delete ship to incompleted.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select list ship to below.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
        }

        private void cmbShiptoMS_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbShiptoMS.SelectedIndex != 0)
            {
                string strCmd = "trim(i_dl_cd) = '" + cmbShiptoMS.Text.ToString().Trim() + "'";
                DataRow[] dr = dsShiptoCD.Tables[0].Select(strCmd);
                if (dr.Count() > 0)
                {
                    tbxShiptoDesc.Text = dr[0]["i_dl_arg_desc"].ToString();

                    strShiptoCD = cmbShiptoMS.Text.ToString().Trim();
                    strShiptoDesc = dr[0]["i_dl_arg_desc"].ToString();
                }
            }
        }

        private bool isHasInListShipto(string shiptoCD)
        {
            string strCmd = "trim(i_shipto_cd) = trim('" + shiptoCD + "') ";
            DataRow[] dr = dsShiptoMS.Tables[0].Select(strCmd);
            if (dr.Count() > 0)
                return true;
            else return false;
        }

        private int addNewShiptoCD(string groupCD, string groupDesc, string shiptoCD, string shiptoDesc)
        {
            try
            {
                int result = shiptoms.insertShiptoMS(groupCD, groupDesc, shiptoCD, shiptoDesc);
                return result;
            }
            catch
            {

                return -1;
            }
        }

        private void dgvListShipto_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (procressCls == "del")
            {
                string shiptoCD = dgvListShipto.Rows[e.RowIndex].Cells["i_shipto_cd"].Value.ToString();
                string shiptoDesc = dgvListShipto.Rows[e.RowIndex].Cells["i_shipto_desc"].Value.ToString();

                strShiptoCD = shiptoCD;
                strShiptoDesc = shiptoDesc;

                tbxShiptoCD.Text = shiptoCD;
                tbxShiptoDesc.Text = shiptoDesc;
            }
        }

        

    }
}
