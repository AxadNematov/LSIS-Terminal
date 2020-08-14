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

namespace WindowsFormsApp1
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
            InitGrid();
            label1.Parent = this;
            label1.BackColor = Color.Transparent;
        }

        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        string app_dir = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString());
        string app_dir_temp = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString()) + "\\Temp\\";
        string show_number = "";

        public void InitGrid()
        {
            dataGridView2.RowCount = 0;
            dataGridView2.ColumnCount = 2;
            dataGridView2.RowHeadersWidth = 30;
            //
            dataGridView2.Columns[0].HeaderText = "#";
            dataGridView2.Columns[1].HeaderText = "ID коробки";
            //
            dataGridView2.Columns[0].Width = 50;
            dataGridView2.Columns[1].Width = 90;
        }

        public string CheckOrderExists(string order_id)
        {
            string oid = "";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT id_order FROM orders WHERE id_order = '" + order_id + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            oid = reader.GetValue(0).ToString();
                        }
                    }
                }
                connection.Close();
            }
            return oid;
        }

        public void Attach_Black_Grey_ID(string black_id, string grey_id)
        {
            //black_id = "74663";
            //grey_id = "123456789012";
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE Boxes SET black_id='" + grey_id + "', grey_id='" + grey_id + "' WHERE black_id='" + black_id + "'";
            command.ExecuteNonQuery();
            //
            command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE items_in_boxes SET black_id='" + grey_id + "' WHERE black_id='" + black_id + "'";
            command.ExecuteNonQuery();
            //
            string box_numb = "";
            string id_order = "";
            string pl_id = "";

            using (connection = new SqlConnection(conString))
            {
                connection.Open();
                using (command = new SqlCommand("SELECT box_numb, id_order, pl_id FROM items_in_boxes WHERE black_id = '" + grey_id + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            box_numb = reader.GetValue(0).ToString();
                            id_order = reader.GetValue(0).ToString();
                            pl_id = reader.GetValue(0).ToString();
                        }
                    }
                }
                connection.Close();
            }
            //
            MessageBox.Show("Коробка номер " + box_numb + "из Packing List  " + pl_id + "  от заказа (PO)  " + id_order + " успешно связана с глобальным ID коробки на складе.", "Сообщение",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button8_attach_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (textBox10.Text.Length == 12) {
                string res = CheckOrderExists(textBox10.Text);
                if (res != "")
                {
                    if (checkBox1.Checked == true) {
                        groupBox1.Enabled = true;
                        checkBox1.Enabled = true;
                        checkBox2.Enabled = true;
                        button9_attach_done.Enabled = true;
                        textBox11.Focus();
                    }
                    if (checkBox2.Checked == true)
                    {
                        groupBox2.Enabled = true;
                        checkBox1.Enabled = true;
                        checkBox2.Enabled = true;
                        button2.Enabled = true;
                        textBox3.Focus();
                    }
                }
                else
                {
                    textBox11.Clear();
                    label_boxn.Text = "- - - - -";
                    label_grey_id.Text = "- - - - - - - - - - - -";
                    label4.Text = "- - - - - - - - - - - -";
                    //
                    textBox3.Clear();
                    dataGridView2.RowCount = 0;
                    textBox10.Focus();
                    groupBox1.Enabled = false;
                    groupBox2.Enabled = false;
                    checkBox1.Enabled = false;
                    checkBox2.Enabled = false;
                    MessageBox.Show("Заказа (PO) с указанным ID не существует.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }else {
                button9_attach_done.Enabled = false;
                textBox11.Clear();
                label_boxn.Text = "- - - - -";
                label_grey_id.Text = "- - - - - - - - - - - -";
                label4.Text = "- - - - - - - - - - - -";

                //
                textBox3.Clear();
                groupBox1.Enabled = false;
                groupBox2.Enabled = false;
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                //
                textBox10.Focus();
                dataGridView2.RowCount = 0;
            }
        }

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        private void timer_scan_Tick(object sender, EventArgs e)
        {
            //MessageBox.Show("Test.", "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Information);
            timer_scan.Enabled = false;
            string str = textBox11.Text;
            if (str.Length == 5) {
                label_boxn.Text = str;
            }
            if (str.Length == 12) {
                label_grey_id.Text = str;
            }
            //
            if (show_number == "") { show_number = str; }
            label1.Visible = true;
            label1.Text = show_number;
            timer_show_number.Enabled = true;
            //
            textBox11.Clear();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            timer_scan.Enabled = true;
        }

        private void label30_Click(object sender, EventArgs e)
        {
            textBox10.Clear();
        }

        private void button9_attach_done_Click(object sender, EventArgs e)
        {
            if (label_boxn.Text != "- - - - - -" && label_grey_id.Text != "- - - - - - - - - - - -")
            {
                Attach_Black_Grey_ID(label_boxn.Text, label_grey_id.Text);
            }
            else
            {
                MessageBox.Show("Необходимо отсканировать ID на коробке и в Packing List.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //textBox10.Clear();
            textBox11.Clear();
            textBox11.Focus();
            label_boxn.Text = "- - - - -";
            label_grey_id.Text = "- - - - - - - - - - - -";
            //button9_attach_done.Enabled = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox11.Clear();
            label_boxn.Text = "- - - - -";
            label_grey_id.Text = "- - - - - - - - - - - -";
            textBox11.Focus();
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            checkBox2.Checked = false;
            //
            groupBox1.Enabled = true;
            groupBox2.Enabled = false;
            textBox11.Focus();
        }

        private void checkBox2_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            //
            groupBox1.Enabled = false;
            groupBox2.Enabled = true;
            textBox3.Focus();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            timer_scan2.Enabled = true;
        }

        private void timer_scan2_Tick(object sender, EventArgs e)
        {
            //MessageBox.Show("Test.", "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Information);
            timer_scan2.Enabled = false;
            //
            string str = textBox3.Text;
            if (str.Length == 5)
            {
                bool check = true;
                for(int i =0; i<dataGridView2.RowCount; i++)
                {
                    if(dataGridView2.Rows[i].Cells[1].Value.ToString()==str)
                    {
                        check = false;
                    }
                }

                if(check==true)
                {
                    //label4.Text = str;
                    dataGridView2.RowCount = dataGridView2.RowCount + 1;
                    dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[0].Value = dataGridView2.RowCount;
                    dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[1].Value = str;
                    if (dataGridView2.RowCount % 2 == 0)
                    {
                        dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[0].Style.BackColor = Color.WhiteSmoke;
                        dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    }
                    else
                    {
                        dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[0].Style.BackColor = Color.White;
                        dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[1].Style.BackColor = Color.White;
                    }
                }
            }
            if (str.Length == 12)
            {
                label4.Text = str;
            }
            //
            if (show_number == "") { show_number = str; }
            label1.Visible = true;
            label1.Text = show_number;
            timer_show_number.Enabled = true;
            //
            textBox3.Clear();
            textBox3.Focus();
        }

        private void dataGridView2_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            textBox3.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
            label4.Text = "- - - - - - - - - - - -";
            dataGridView2.RowCount = 0;
            textBox3.Focus();
        }

        private void textBox10_KeyUp(object sender, KeyEventArgs e)
        {
            //if (e.KeyData == Keys.Tab)
            //{
            //    textBox10.Focus();
            //}
        }

        private void timer_show_number_Tick(object sender, EventArgs e)
        {
            timer_show_number.Enabled = false;
            label1.Visible = false;
            show_number = "";
        }
        public class ListBlackID
        {
            public string blackId { get; set; }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount > 0 && label4.Text != "- - - - - - - - - - - -")
            {
                string poleta = label4.Text;
                List<ListBlackID> listBlackId = new List<ListBlackID>();
                for(int i = 0; i<dataGridView2.Rows.Count; i++)
                {
                    listBlackId.Add(new ListBlackID()
                    {
                        blackId = dataGridView2.Rows[i].Cells[1].Value.ToString()
                    });
                }
                linkPoleta(listBlackId, poleta);
            }
            else
            {
                MessageBox.Show("Необходимо отсканировать ID палеты и ID в Packing List.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            textBox3.Focus();
        }

        public void linkPoleta(List<ListBlackID> list, string idPoleta)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            int i = 0;
            for (i = 0; i < list.Count; i++)
            {
                string blackId;
                string greyId = "";
                bool checkGreyId = false;


                while (checkGreyId == false)
                {
                    greyId = generateGreyId();
                    SqlCommand check_Grey_Id = new SqlCommand("SELECT COUNT(*) FROM [Boxes] WHERE ([grey_id] = @greyid)", connection);
                    check_Grey_Id.Parameters.AddWithValue("@greyid", greyId);
                    int idExist = (int)check_Grey_Id.ExecuteScalar();

                    if (idExist > 0)
                    {
                        checkGreyId = false;
                    }
                    else
                    {
                        checkGreyId = true;
                    }
                }

                blackId = list[i].blackId;

                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE Boxes SET black_id=@greyid, grey_id=@greyid, red_id=@poleta WHERE black_id=@blackid";
                command.Parameters.AddWithValue("@poleta", idPoleta);
                command.Parameters.AddWithValue("@greyid", greyId);
                command.Parameters.AddWithValue("@blackid", blackId);
                command.ExecuteNonQuery();

                SqlCommand command1 = new SqlCommand();
                command1.Connection = connection;
                command1.CommandType = CommandType.Text;
                command1.CommandText = "UPDATE items_in_boxes SET black_id=@greyid WHERE black_id=@blackid";
                command1.Parameters.AddWithValue("@greyid", greyId);
                command1.Parameters.AddWithValue("@blackid", blackId);
                command1.ExecuteNonQuery();
            }

            connection.Close();
            MessageBox.Show("Done!", "Done",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public string generateGreyId()
        {
            string grayId;
            Random rnd = new Random();
            int valueFirst = rnd.Next(999, 9999);
            int valueSecond = rnd.Next(999, 9999);
            int valueThird = rnd.Next(999, 9999);
            grayId = valueFirst.ToString() + valueSecond.ToString() + valueThird.ToString();
            return grayId;
        }
    }
}
