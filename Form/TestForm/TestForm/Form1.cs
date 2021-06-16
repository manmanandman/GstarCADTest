using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace TestForm
{
    public partial class Form1 : Form
    {
        String conn_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\Man\\Desktop\\C_\\manTest\\XdataTest.accdb;Persist Security Info= False";
        String error_msg = "";
        String q = "INSERT INTO XDATA_TEST VALUES ('1001','AUTOPOST')";
        OleDbConnection conn = null;
        
        public Form1()
        {
            InitializeComponent();
        }

        // Connect to database
        private void connnectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                conn = new OleDbConnection(conn_string);
                conn.Open();
                MessageBox.Show("Connected to database");
                queryToolStripMenuItem_Click(sender, e);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // Query all data in database
        private void queryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                error_msg = "";
                q = "SELECT * FROM XDATA_TEST";

                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;

                OleDbDataAdapter a = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();

                a.SelectCommand = command;
                a.Fill(dt);

                result.DataSource = dt;
                result.AutoResizeColumns();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        // Insert mock data in database
        private void addMockingDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                error_msg = "";
                q = "INSERT INTO XDATA_TEST VALUES ('1001','AUTOPOST')";

                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;
                command.ExecuteNonQuery();
                MessageBox.Show("Inserted to database");
                queryToolStripMenuItem_Click(sender,e);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
           
        }

        private void exportPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReportViewer report = new ReportViewer();
            report.Show();
        }

        private void clearDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                error_msg = "";
                q = "DELETE * FROM XDATA_TEST";

                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;

                OleDbDataAdapter a = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();

                a.SelectCommand = command;
                a.Fill(dt);

                MessageBox.Show("Clear database success");
                queryToolStripMenuItem_Click(sender, e);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
    }
}
