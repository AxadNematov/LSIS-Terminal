using System;
using System.IO;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelLibrary.BinaryDrawingFormat;
using ExcelLibrary.BinaryFileFormat;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SqlServer.Server;
using System.Windows;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Runtime.InteropServices;
using System.Threading;
using OfficeOpenXml;


namespace WindowsFormsApp1
{
    public partial class Form6 : Form
    {
        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        private Form1 Frm1;

        public int editType = 0;
        public Form6(Form1 f1)
        {
            InitializeComponent();
            InitGrid();
            Frm1 = f1;
            RefreshList();
        }

        private void Form6_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddCompanyDetails();
        }
        public void AddCompanyDetails()
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO company_details(company, address, contact, port_name, final_destination) VALUES(@company, @address, @contact, @port, @fdestination)";
            command1.Parameters.AddWithValue("@company", textBox1.Text);
            command1.Parameters.AddWithValue("@address", textBox2.Text);
            command1.Parameters.AddWithValue("@contact", textBox3.Text);
            command1.Parameters.AddWithValue("@port", textBox4.Text);
            command1.Parameters.AddWithValue("@fdestination", textBox5.Text);
            command1.ExecuteNonQuery();

            MessageBox.Show("New company details succesfully added!", "Done",
            MessageBoxButtons.OK, MessageBoxIcon.Information);

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            RefreshList();
        }
        public void InitGrid()
        {
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 6;

            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "Company";
            dataGridView1.Columns[2].HeaderText = "Address";
            dataGridView1.Columns[3].HeaderText = "Contact";
            dataGridView1.Columns[4].HeaderText = "Port";
            dataGridView1.Columns[5].HeaderText = "Final Destination";

            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[5].Width = 200;
        }
        public class CompanyDetails
        {
            public string company { get; set; }
            public string address { get; set; }
            public string contact { get; set; }
            public string port { get; set; }
            public string finalDestination { get; set; }
        }

        public void RefreshList()
        {
            List<CompanyDetails> details_list = new List<CompanyDetails>();
            string sql = "SELECT company, address, contact, port_name, final_destination FROM company_details ORDER BY company";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CompanyDetails company = new CompanyDetails();
                            company.company = reader.GetValue(0).ToString();
                            company.address = reader.GetValue(1).ToString();
                            company.contact = reader.GetValue(2).ToString();
                            company.port = reader.GetValue(3).ToString();
                            company.finalDestination = reader.GetValue(4).ToString();
                            details_list.Add(company);
                        }
                    }
                }
                connection.Close();
            }
            dataGridView1.RowCount = details_list.Count;
            for(int i = 0; i<details_list.Count; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = i + 1;
                dataGridView1.Rows[i].Cells[1].Value = details_list[i].company;
                dataGridView1.Rows[i].Cells[2].Value = details_list[i].address;
                dataGridView1.Rows[i].Cells[3].Value = details_list[i].contact;
                dataGridView1.Rows[i].Cells[4].Value = details_list[i].port;
                dataGridView1.Rows[i].Cells[5].Value = details_list[i].finalDestination;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int selected = dataGridView1.CurrentCell.RowIndex;
            textBox10.Text = dataGridView1.Rows[selected].Cells[1].Value.ToString();
            textBox9.Text = dataGridView1.Rows[selected].Cells[2].Value.ToString();
            textBox8.Text = dataGridView1.Rows[selected].Cells[3].Value.ToString();
            textBox7.Text = dataGridView1.Rows[selected].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[selected].Cells[5].Value.ToString();
            editType = 1;
            textBox10.ReadOnly = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            saveChanges();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int selected = dataGridView1.CurrentCell.RowIndex;
            textBox10.Text = dataGridView1.Rows[selected].Cells[1].Value.ToString();
            textBox9.Text = dataGridView1.Rows[selected].Cells[2].Value.ToString();
            textBox8.Text = dataGridView1.Rows[selected].Cells[3].Value.ToString();
            textBox7.Text = dataGridView1.Rows[selected].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[selected].Cells[5].Value.ToString();
            editType = 2;
            textBox10.ReadOnly = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox10.ReadOnly = false;
            editType = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            deleteCompanyDetails();
        }

        public void saveChanges()
        {
            if (editType == 1)
            {
                int selected = dataGridView1.CurrentCell.RowIndex;
                string company = dataGridView1.Rows[selected].Cells[1].Value.ToString();
                string address = dataGridView1.Rows[selected].Cells[2].Value.ToString();
                string contact = dataGridView1.Rows[selected].Cells[3].Value.ToString();
                string port = dataGridView1.Rows[selected].Cells[4].Value.ToString();
                string fDestination = dataGridView1.Rows[selected].Cells[5].Value.ToString();

                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE company_details SET company='" + textBox10.Text + "', address='" + textBox9.Text + "', contact='" + textBox8.Text + "', port_name='" + textBox7.Text + "', final_destination='" + textBox6.Text + "' WHERE company='" + company + "' and address='" + address + "' and contact='" + contact + "' and port_name='" + port + "' and final_destination='" + fDestination + "'";
                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Changes succesfully saved!", "Done",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (editType == 2)
            {
                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                SqlCommand command1 = new SqlCommand();
                command1.Connection = connection;
                command1.CommandType = CommandType.Text;
                command1.CommandText = "INSERT INTO company_details(company, address, contact, port_name, final_destination) VALUES(@company, @address, @contact, @port, @fdestination)";
                command1.Parameters.AddWithValue("@company", textBox10.Text);
                command1.Parameters.AddWithValue("@address", textBox9.Text);
                command1.Parameters.AddWithValue("@contact", textBox8.Text);
                command1.Parameters.AddWithValue("@port", textBox7.Text);
                command1.Parameters.AddWithValue("@fdestination", textBox6.Text);
                command1.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("New version succesfully saved!", "Done",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox10.ReadOnly = false;
            }
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            editType = 0;
            RefreshList();
        }
        public void deleteCompanyDetails()
        {
            int selected = dataGridView1.CurrentCell.RowIndex;
            string company = dataGridView1.Rows[selected].Cells[1].Value.ToString();
            string address = dataGridView1.Rows[selected].Cells[2].Value.ToString();
            string contact = dataGridView1.Rows[selected].Cells[3].Value.ToString();
            string port = dataGridView1.Rows[selected].Cells[4].Value.ToString();
            string fDestination = dataGridView1.Rows[selected].Cells[5].Value.ToString();

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "DELETE FROM company_details WHERE company='" + company + "' and address='" + address + "' and contact='" + contact + "' and port_name='" + port + "' and final_destination='" + fDestination + "'";
            command.ExecuteNonQuery();
            connection.Close();

            MessageBox.Show("Deleted!", "Done",
            MessageBoxButtons.OK, MessageBoxIcon.Information);

            RefreshList();
        }
    }
}
