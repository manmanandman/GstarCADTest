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

namespace test00
{
    public partial class Form1 : Form
    {

        String conn_string = HelloCmd.conn_string;
        OleDbConnection conn = null;
        String q = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void connectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                conn = new OleDbConnection(conn_string);
                conn.Open();
                MessageBox.Show("Connected to database");
                QueryDatabase();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void QueryDatabase()
        {
            try
            {
                string q = "SELECT * FROM XDATA_TEST";

                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;

                OleDbDataAdapter a = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();

                a.SelectCommand = command;
                a.Fill(dt);

                Result.DataSource = dt;
                Result.AutoResizeColumns();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                conn = new OleDbConnection(conn_string);
                conn.Open();
                QueryDatabase();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void exportToPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReportViewer report = new ReportViewer();
            report.Show();
        }

        private void clearDataInDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                q = "DELETE * FROM XDATA_TEST";

                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;

                OleDbDataAdapter a = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();

                a.SelectCommand = command;
                a.Fill(dt);

                MessageBox.Show("Clear database success");
                QueryDatabase();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void disconnectToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void refreshDatabseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string q = "SELECT * FROM XDATA_TEST";

                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;

                OleDbDataAdapter a = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();

                a.SelectCommand = command;
                a.Fill(dt);

                Result.DataSource = dt;
                Result.AutoResizeColumns();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void Result_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
