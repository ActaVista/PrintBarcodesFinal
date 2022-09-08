
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using Microsoft.VisualBasic;

namespace PrintBarcodesFromGridview
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        string cs = "server=actavista.net;uid=barcodeadmin;pwd=PA$$word@421640;database=barcodedb;";
        string currency = "¥";

        MySqlConnection con;
        MySqlCommand cmd;
        MySqlDataReader rdr;
        DataTable DTable;

        private DataTable dt = new DataTable("Products");
        private List<Image> barcodes = new List<Image>();

        private void Reset()
        {
            txtSearchProduct.Text = "";
            lvPrint.Items.Clear();
            dgvSelect.Rows.Clear();
            nudExclude.Value = 0;
            chkShowPrices.Checked = false;
            if (DTable != null) { DTable.Rows.Clear(); }
            if (dt != null) { dt.Rows.Clear(); }
            GetData();
            txtSearchProduct.Select();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Reset();
        }

        private void GetData()
        {
            try
            {
                con = new MySqlConnection(cs);
                con.Open();
                cmd = new MySqlCommand("Select RTRIM(Product_Code),RTRIM(Product_Name),RTRIM(JanCode1),RTRIM(P_PriceRate) From Product_New Order By Product_Name", con);
                rdr = cmd.ExecuteReader();
                dgvSelect.Rows.Clear();
                while (rdr.Read())
                {
                    dgvSelect.Rows.Add(rdr[0], rdr[1], rdr[2].ToString().Replace(":", ""), rdr[3]);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            if (lvPrint.Items.Count <= 0)
            {
                MessageBox.Show("There is no product in listview to generate barcode ", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (dt.Rows.Count <= 0 || DTable.Rows.Count <= 0)
            {
                MessageBox.Show("There is no product in listview to generate barcode ", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                Cursor= Cursors.WaitCursor; 
                if (nudExclude.Value > 0)
                {
                    for (int i = 0; i < nudExclude.Value; i++)
                    {
                        DataRow dtdataRow = dt.NewRow();
                        dt.Rows.InsertAt(dtdataRow, i);
                        dt.AcceptChanges();
                    }
                    for (int i = 0; i < nudExclude.Value; i++)
                    {
                        DataRow dtableRow = DTable.NewRow();
                        DTable.Rows.InsertAt(dtableRow, i);
                        DTable.AcceptChanges();
                    }
                }
                foreach (ListViewItem item in lvPrint.Items)
                {
                    if (!item.Checked)
                    {
                        //Remove unchecked Items
                        lvPrint.Items.Remove(item);
                        var rows = DTable.Select("MRP = '" + item.Text.Trim() + "'");
                        foreach (var row in rows)
                        { row.Delete(); }
                        DTable.AcceptChanges();
                        var irows = dt.Select("MRP = '" + item.Text.Trim() + "'");
                        foreach (var row in irows)
                        { row.Delete(); }
                        dt.AcceptChanges();
                    }
                }
                if (DTable.Rows.Count > 0)
                {
                    if (chkShowPrices.Checked)
                    {
                        DTable = dt.Copy();
                        DTable.WriteXmlSchema("BarcodeLabelPrinting1.xml");
                        rptBarcodeLabelPrinting rpt = new rptBarcodeLabelPrinting();
                        rpt.SetDataSource(DTable);
                        rpt.SetParameterValue("p1",currency);
                        frmReport frmReport = new frmReport();
                        frmReport.CrystalReportViewer1.ReportSource = rpt;
                        frmReport.ShowDialog();
                        rpt.Close();
                        rpt.Dispose();
                    }
                    else
                    {
                        for (int i = 0; i < DTable.Rows.Count; i++)
                        { DTable.Rows[i][3] = ""; }
                        DTable.WriteXmlSchema("BarcodeLabelPrinting1.xml");
                        rptBarcodeLabelPrinting rpt = new rptBarcodeLabelPrinting();
                        rpt.SetDataSource(DTable);
                        rpt.SetParameterValue("p1", "");
                        frmReport frmReport = new frmReport();
                        frmReport.CrystalReportViewer1.ReportSource = rpt;
                        frmReport.ShowDialog();
                        rpt.Close();
                        rpt.Dispose();
                    }
                    Reset();
                    timer1.Start();
                    MessageBox.Show("Task Completed", "Process Complted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                { MessageBox.Show("Looks like there's nothing to write buddy", "Let's try it again", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            Cursor = DefaultCursor;
            timer1.Stop();
        }

        private void txtSearchProduct_TextChanged(object sender, EventArgs e)
        {
            try
            {
                con = new MySqlConnection(cs);
                con.Open();
                cmd = new MySqlCommand("Select RTRIM(Product_Code),RTRIM(Product_Name),RTRIM(JanCode1),RTRIM(P_PriceRate) From Product_New Where Product_Code Like @SearchString Or Product_Name Like @SearchString Order By Product_Name", con);
                cmd.Parameters.AddWithValue("@SearchString", "%" + txtSearchProduct.Text.Trim() + "%");
                rdr = cmd.ExecuteReader();
                dgvSelect.Rows.Clear();
                while (rdr.Read())
                {
                    dgvSelect.Rows.Add(rdr[0], rdr[1], rdr[2].ToString().Replace(":", ""), rdr[3]);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgvSelect_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            bool isExists = false;
            int Qty;
            foreach (ListViewItem viewItem in lvPrint.Items)
                if (viewItem.Text == dgvSelect.Rows[e.RowIndex].Cells[0].Value.ToString())
                {
                    isExists = true;
                    break;
                }

            if (isExists) { MessageBox.Show("This item has already been selected buddy", "Let's try something new", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }


            string PrintQty = Interaction.InputBox("How many would you like to print?", "0");
            if (int.TryParse(PrintQty, out Qty))
            {
                var item = new ListViewItem();
                item.Text = dgvSelect.Rows[e.RowIndex].Cells[0].Value.ToString();
                item.SubItems.Add(dgvSelect.Rows[e.RowIndex].Cells[1].Value.ToString());
                item.SubItems.Add(dgvSelect.Rows[e.RowIndex].Cells[2].Value.ToString());
                item.SubItems.Add(dgvSelect.Rows[e.RowIndex].Cells[3].Value.ToString());
                item.SubItems.Add(PrintQty.Trim());
                lvPrint.Items.Add(item);

                item.Checked = true;

                if (dt.Columns.Count <= 0)
                {
                    dt.Columns.Add("MRP", typeof(string));
                    dt.Columns.Add("ProductName", typeof(string));
                    dt.Columns.Add("Barcode", typeof(string));
                    dt.Columns.Add("SalesRate", typeof(string));
                }
                for (int i = 0; i < Convert.ToInt32(PrintQty); i++)
                {
                    dt.Rows.Add(dgvSelect.Rows[e.RowIndex].Cells[0].Value.ToString(), dgvSelect.Rows[e.RowIndex].Cells[1].Value.ToString(), dgvSelect.Rows[e.RowIndex].Cells[2].Value.ToString(),string.Format("{0}円", dgvSelect.Rows[e.RowIndex].Cells[3].Value.ToString()));
                }
                DataSet ds = new DataSet();
                ds.Tables.Add(dt.Copy());
                DTable = new DataTable();
                DTable = ds.Tables[0];
            }
            else
            { return; }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Reset();
        }
    }
}
