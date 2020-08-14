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
    public partial class Form2 : Form
    {
        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        string app_dir = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString());
        string app_dir_temp = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString()) + "\\Temp\\";


        private Form1 form1;
        public Form2(Form1 f1)
        {
            InitializeComponent();
            InitDataGridView();
            form1 = f1;
        }
        
        public string TextBox_Z
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        public string TextBox_X
        {
            get { return textBox3.Text; }
            set { textBox3.Text = value; }
        }

        public string TextBox_Y
        {
            get { return textBox4.Text; }
            set { textBox4.Text = value; }
        }

        public string TextBox_Z_nloc
        {
            get { return textBox7.Text; }
            set { textBox7.Text = value; }
        }

        public string TextBox_X_nloc
        {
            get { return textBox6.Text; }
            set { textBox6.Text = value; }
        }

        public string TextBox_Y_nloc
        {
            get { return textBox5.Text; }
            set { textBox5.Text = value; }
        }

        public DataGridView Frm2_DataGridView1
        {
            get { return dataGridView1; }
            set { dataGridView1 = value; }
        }

        public class Boxes
        {
            public string boxNumber { get; set; }
            public string greyId { get; set; }
            public string location_x { get; set; }
            public string location_y { get; set; }
            public string location_z { get; set; }
            public string location_b { get; set; }
        }
        
        public class BoxItems
        {
            public string boxNumber { get; set; }
            public string partName { get; set; }
            public string partCode { get; set; }
            public string amount { get; set; }
            public string netWeigh { get; set; }
            public string grossWeight { get; set; }
        }

        public void InitDataGridView()
        {
            
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 7;

            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "# Коробки";
            dataGridView1.Columns[2].HeaderText = "ID Коробки";
            dataGridView1.Columns[3].HeaderText = "Секция";
            dataGridView1.Columns[4].HeaderText = "Ярус";
            dataGridView1.Columns[5].HeaderText = "Стелаж";
            dataGridView1.Columns[6].HeaderText = "Блок";

            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 90;
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].Width = 70;
            dataGridView1.Columns[4].Width = 70;
            dataGridView1.Columns[5].Width = 70;
            dataGridView1.Columns[6].Width = 70;
            //
            dataGridView2.RowCount = 0;
            dataGridView2.ColumnCount = 7;

            dataGridView2.Columns[0].HeaderText = "#";
            dataGridView2.Columns[1].HeaderText = "# Коробки";
            dataGridView2.Columns[2].HeaderText = "Название продукта";
            dataGridView2.Columns[3].HeaderText = "Код продукта";
            dataGridView2.Columns[4].HeaderText = "Кол-во";
            dataGridView2.Columns[5].HeaderText = "NET";
            dataGridView2.Columns[6].HeaderText = "GROSS";

            dataGridView2.Columns[0].Width = 40;
            dataGridView2.Columns[1].Width = 90;
            dataGridView2.Columns[2].Width = 220;
            dataGridView2.Columns[3].Width = 100;
            dataGridView2.Columns[4].Width = 70;
            dataGridView2.Columns[5].Width = 70;
            dataGridView2.Columns[6].Width = 70;
        }

        private void button8_search_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = 0;
            int selectedIndex = comboBox1.SelectedIndex;
            if (selectedIndex == 3)
            {
                string x = "";
                string y = "";
                string z = "";
                string b = "";
                x = textBox3.Text;
                y = textBox4.Text;
                z = textBox2.Text;
                b = comboBox3.Text;
                if (b != "") {
                    constructDataGridByCoordinates(x, y, z, b);
                } else { 
                    constructDataGridByCoordinates2(x, y, z);
                }
            }
            else
            {
                string searchArgument = "";
                searchArgument = textBox1.Text;
                constructDataGrid(searchArgument, selectedIndex);
            }
        }

        public void constructDataGrid(string search, int index)
        {
            List<Boxes> boxes = new List<Boxes>();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            string connectionString = "SELECT box_number, grey_id, location_x, location_y, location_z, block FROM Boxes WHERE grey_id!='---' ORDER BY box_number";
            switch (index)
            {
                case 0:
                    connectionString = "SELECT box_number, grey_id, location_x, location_y, location_z, block FROM Boxes WHERE grey_id!='---' ORDER BY box_number";
                    break;
                case 1:
                    connectionString = "SELECT box_number, grey_id, location_x, location_y, location_z, block FROM Boxes WHERE grey_id=@search ORDER BY box_number";
                    break;
                case 2:
                    connectionString = "select distinct Boxes.box_number, Boxes.grey_id, Boxes.location_x, Boxes.location_y, block Boxes.location_z from Boxes join items_in_boxes on Boxes.black_id = items_in_boxes.black_id where items_in_boxes.id_order = @search and Boxes.grey_id!='---'";
                    break;
            }
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = connectionString;
            comm.Parameters.AddWithValue("@search", search);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                boxes.Add(new Boxes()
                {
                    boxNumber = Convert.ToString(reader["box_number"]),
                    greyId = Convert.ToString(reader["grey_id"]),
                    location_x = Convert.ToString(reader["location_x"]),
                    location_y = Convert.ToString(reader["location_y"]),
                    location_z = Convert.ToString(reader["location_z"]),
                    location_b = Convert.ToString(reader["block"])
                });
            }
            fillDataGridView(boxes);
        }

        public void constructDataGridByCoordinates(string x, string y, string z, string b)
        {
            List<Boxes> boxes = new List<Boxes>();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            string connectionString = "SELECT box_number, grey_id, location_x, location_y, location_z, block FROM Boxes WHERE location_x=@corx and location_y=@cory and location_z=@corz and block=@corb";
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = connectionString;
            comm.Parameters.AddWithValue("@corx", x);
            comm.Parameters.AddWithValue("@cory", y);
            comm.Parameters.AddWithValue("@corz", z);
            comm.Parameters.AddWithValue("@corb", b);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                boxes.Add(new Boxes()
                {
                    boxNumber = Convert.ToString(reader["box_number"]),
                    greyId = Convert.ToString(reader["grey_id"]),
                    location_x = Convert.ToString(reader["location_x"]),
                    location_y = Convert.ToString(reader["location_y"]),
                    location_z = Convert.ToString(reader["location_z"]),
                    location_b = Convert.ToString(reader["block"])
                });
            }
            fillDataGridView(boxes);
        }

        public void constructDataGridByCoordinates2(string x, string y, string z)
        {
            List<Boxes> boxes = new List<Boxes>();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            string connectionString = "SELECT box_number, grey_id, location_x, location_y, location_z, block FROM Boxes WHERE location_x=@corx and location_y=@cory and location_z=@corz";
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = connectionString;
            comm.Parameters.AddWithValue("@corx", x);
            comm.Parameters.AddWithValue("@cory", y);
            comm.Parameters.AddWithValue("@corz", z);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                boxes.Add(new Boxes()
                {
                    boxNumber = Convert.ToString(reader["box_number"]),
                    greyId = Convert.ToString(reader["grey_id"]),
                    location_x = Convert.ToString(reader["location_x"]),
                    location_y = Convert.ToString(reader["location_y"]),
                    location_z = Convert.ToString(reader["location_z"]),
                    location_b = Convert.ToString(reader["block"])
                });
            }
            fillDataGridView(boxes);
        }

        public void fillDataGridView(List<Boxes> boxList)
        {
            dataGridView1.RowCount = boxList.Count;
            int i = 0;
            for (i = 0; i < boxList.Count; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = (i + 1).ToString();
                dataGridView1.Rows[i].Cells[1].Value = boxList[i].boxNumber;
                dataGridView1.Rows[i].Cells[2].Value = boxList[i].greyId;
                dataGridView1.Rows[i].Cells[3].Value = boxList[i].location_x;
                dataGridView1.Rows[i].Cells[4].Value = boxList[i].location_y;
                dataGridView1.Rows[i].Cells[5].Value = boxList[i].location_z;
                dataGridView1.Rows[i].Cells[6].Value = boxList[i].location_b;
                if (i % 2 == 0)
                {
                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[6].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[6].Style.BackColor = Color.White;
                }
            }
            dataGridView2.RowCount = 0;
            //dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.None;
        }

        public void selectBoxItems(string grey_id)
        {
            List<BoxItems> items = new List<BoxItems>();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "select items_in_boxes.box_numb, items_in_boxes.part_name, items_in_boxes.part_code, items_in_boxes.amount, items_in_boxes.net, items_in_boxes.gross from items_in_boxes join Boxes on items_in_boxes.black_id = Boxes.black_id where Boxes.grey_id = @greyId";
            comm.Parameters.AddWithValue("@greyId", grey_id);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                items.Add(new BoxItems()
                {
                    boxNumber = Convert.ToString(reader["box_numb"]),
                    partName = Convert.ToString(reader["part_name"]),
                    partCode = Convert.ToString(reader["part_code"]),
                    amount = Convert.ToString(reader["amount"]),
                    netWeigh = Convert.ToString(reader["net"]),
                    grossWeight = Convert.ToString(reader["gross"])
                });
            }
            fillBoxItems(items);
        }

        public void fillBoxItems(List<BoxItems> boxItems)
        {
            int total_amount = 0;
            int total_net = 0;
            int total_gross = 0;
            dataGridView2.RowCount = boxItems.Count;
            if (boxItems.Count > 0) {
                dataGridView2.RowCount = boxItems.Count + 1;
            }
            int i = 0;
            for (i = 0; i < boxItems.Count; i++)
            {
                total_amount = total_amount + Convert.ToInt32(boxItems[i].amount);
                total_net = total_net + Convert.ToInt32(boxItems[i].netWeigh);
                total_gross = total_gross + Convert.ToInt32(boxItems[i].grossWeight);
                //
                dataGridView2.Rows[i].Cells[0].Value = (i + 1).ToString();
                dataGridView2.Rows[i].Cells[1].Value = boxItems[i].boxNumber;
                dataGridView2.Rows[i].Cells[2].Value = boxItems[i].partName;
                dataGridView2.Rows[i].Cells[3].Value = boxItems[i].partCode;
                dataGridView2.Rows[i].Cells[4].Value = boxItems[i].amount;
                dataGridView2.Rows[i].Cells[5].Value = boxItems[i].netWeigh;
                dataGridView2.Rows[i].Cells[6].Value = boxItems[i].grossWeight;
                if (i % 2 == 0)
                {
                    dataGridView2.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView2.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView2.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    dataGridView2.Rows[i].Cells[3].Style.BackColor = Color.WhiteSmoke;
                    dataGridView2.Rows[i].Cells[4].Style.BackColor = Color.WhiteSmoke;
                    dataGridView2.Rows[i].Cells[5].Style.BackColor = Color.WhiteSmoke;
                    dataGridView2.Rows[i].Cells[6].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView2.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView2.Rows[i].Cells[1].Style.BackColor = Color.White;
                    dataGridView2.Rows[i].Cells[2].Style.BackColor = Color.White;
                    dataGridView2.Rows[i].Cells[3].Style.BackColor = Color.White;
                    dataGridView2.Rows[i].Cells[4].Style.BackColor = Color.White;
                    dataGridView2.Rows[i].Cells[5].Style.BackColor = Color.White;
                    dataGridView2.Rows[i].Cells[6].Style.BackColor = Color.White;
                }
            }
            dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[3].Value = "Итого";
            dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[4].Value = total_amount;
            dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[5].Value = total_net;
            dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[6].Value = total_gross;
        }

        public void UpdateBoxLocation(string grey_id, string x, string y, string z, string b)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE Boxes SET location_x='" + x + "', location_y='" + y + "', location_z='" + z + "', block='" + b + "' WHERE grey_id='" + grey_id + "'";
            command.ExecuteNonQuery();
        }

        public void RemoveBoxLocation(string grey_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE Boxes SET location_x='---', location_y='---', location_z='---', block='---' WHERE grey_id='" + grey_id + "'";
            command.ExecuteNonQuery();
        }

        public void AddNewBox(string bg_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "INSERT INTO Boxes(box_number, black_id, grey_id, location_x, location_y, location_z) VALUES(1, @black_id, @grey_id, '---', '---', '---')";
            command.Parameters.AddWithValue("@black_id", bg_id);
            command.Parameters.AddWithValue("@grey_id", bg_id);
            command.ExecuteNonQuery();
        }

        public int BoxExists(string grey_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            //
            SqlCommand check_Grey_Exist = new SqlCommand("SELECT COUNT(Boxes.grey_id) FROM Boxes WHERE grey_id=@greyid", connection);
            check_Grey_Exist.Parameters.AddWithValue("@greyid", grey_id);
            int greyIDExist = (int)check_Grey_Exist.ExecuteScalar();
            //
            return greyIDExist;
        } 

        public void checkDelete(string greyId)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            //
            SqlCommand check_Grey_Exist = new SqlCommand("SELECT COUNT(Boxes.grey_id) FROM Boxes WHERE grey_id=@greyid", connection);
            check_Grey_Exist.Parameters.AddWithValue("@greyid", greyId);
            int greyIDExist = (int)check_Grey_Exist.ExecuteScalar();
            //
            if (greyIDExist > 0)
            {
                SqlCommand check_Grey_Id = new SqlCommand("SELECT COUNT(items_in_boxes.black_id) FROM items_in_boxes join Boxes on items_in_boxes.black_id=Boxes.black_id  WHERE Boxes.grey_id=@greyid", connection);
                check_Grey_Id.Parameters.AddWithValue("@greyid", greyId);
                int idExist = (int)check_Grey_Id.ExecuteScalar();

                if (idExist > 0)
                {
                    MessageBox.Show("Невозможно удалить коробку так как она не пустая.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "DELETE FROM Boxes WHERE grey_id=@greyid";
                    command.Parameters.AddWithValue("@greyid", greyId);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Коробка удалена со склада.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Коробки с указанным ID не существует.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void addItem(string boxNumber, string partName, string partCode, string amount, string net, string gross, string idOrder, string black_id, string pl_id, string pl_date, string invoice_id, string invoice_date)
        {
            //string boxNumber = "";
            //string partName = "";
            //string partCode = "";
            //string amount = "";
            //string net = "";
            //string gross = "";
            //string idOrder = "";
            //string black_id = "";
            //string pl_id = "";
            //string pl_date = "";
            //string invoice_id = "";
            //string invoice_date = "";

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "insert into items_in_boxes(box_numb, part_name, part_code, amount, net, gross, id_order, black_id, pl_id, pl_date, invoice_id, invoice_date) values(@boxno, @partname, @partcode, @amount, @net, @gross, @idorder, @blackid, @plid, @pldate, @invoiceid, @invoicedate)";
            command.Parameters.AddWithValue("@boxno", boxNumber);
            command.Parameters.AddWithValue("@partname", partName);
            command.Parameters.AddWithValue("@partcode", partCode);
            command.Parameters.AddWithValue("@amount", amount);
            command.Parameters.AddWithValue("@net", net);
            command.Parameters.AddWithValue("@gross", gross);
            command.Parameters.AddWithValue("@idorder", idOrder);
            command.Parameters.AddWithValue("@blackid", black_id);
            command.Parameters.AddWithValue("@plid", pl_id);
            command.Parameters.AddWithValue("@pldate", pl_date);
            command.Parameters.AddWithValue("@invoiceid", invoice_id);
            command.Parameters.AddWithValue("@invoicedate", invoice_date);
            command.ExecuteNonQuery();
            connection.Close();
            MessageBox.Show("Продукты успешно добавлены в коробку..", "Сообщение",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 3)
            {
                textBox1.Enabled = false;
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                comboBox3.Enabled = true;
            }
            else
            {
                textBox1.Enabled = true;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                comboBox3.Enabled = false;
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView1.CurrentCell.RowIndex;
                string grey_id = dataGridView1.Rows[selected].Cells[2].Value.ToString();
                textBox8.Text = grey_id;
                textBox9.Text = grey_id;
                selectBoxItems(grey_id);
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int selected;
            try {
                selected = dataGridView1.CurrentCell.RowIndex;
                string x = dataGridView1.Rows[selected].Cells[3].Value.ToString();
                string grey_id = dataGridView1.Rows[selected].Cells[2].Value.ToString();
                //if (x == "---") {
                    // change appearence to box_location mode
                    button1.Enabled = false;
                    button7.Enabled = false;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button8_search.Enabled = false;
                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox7.Clear();
                    textBox5.Enabled = true;
                    textBox6.Enabled = true;
                    textBox7.Enabled = true;
                    comboBox2.Enabled = true;
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    textBox3.Enabled = false;
                    textBox4.Enabled = false;
                    //
                    form1.shelf_mode = "box_location";
                //}
                //else {
                //    MessageBox.Show("Выбранная коробка уже расположена на складе.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}
            }
            catch { }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // change appearence back to search mode
            form1.shelf_mode = "search";
            button1.Enabled = true;
            button7.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;
            button8_search.Enabled = true;
            dataGridView1.Enabled = true;
            dataGridView2.Enabled = true;
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            comboBox2.Enabled = false;
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            //
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView1.CurrentCell.RowIndex;
                string grey_id = dataGridView1.Rows[selected].Cells[2].Value.ToString();
                RemoveBoxLocation(grey_id);
                UpdateBoxLocation(grey_id, textBox6.Text, textBox5.Text, textBox7.Text, comboBox2.Text);
                constructDataGrid("", 0);
                //
                form1.MakeShelfsStructure(Convert.ToInt32(textBox7.Text));
                // change appearence back to search mode
                form1.shelf_mode = "search";
                button1.Enabled = true;
                button7.Enabled = true;
                button2.Enabled = false;
                button3.Enabled = false;
                button8_search.Enabled = true;
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                textBox7.Enabled = false;
                comboBox2.Enabled = false;
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;

                //
            }
            catch { }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView1.CurrentCell.RowIndex;
                string grey_id = dataGridView1.Rows[selected].Cells[2].Value.ToString();
                //
                DialogResult dialogResult = MessageBox.Show("Вы уверены что хотите снять выбранную коробку со стелажа.", "Предупреждение.", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes) {
                    RemoveBoxLocation(grey_id);
                    constructDataGrid("", 0);
                    //do something
                }
                else if (dialogResult == DialogResult.No) {
                    //do something else
                }
                //
                form1.MakeShelfsStructure(form1._current_shelf_Z);
                //form1.MarkShelfs(Convert.ToInt32(textBox7.Text));
                //
            }
            catch { }
        }

        private void label19_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string grey_id = textBox9.Text;
            try {
                int g = Convert.ToInt32(grey_id);
                if (grey_id.Length == 12) {
                    AddNewBox(grey_id);
                    MessageBox.Show("Коробка с ID " + grey_id + "успешно добавлена в систему, но не расположена на складе.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else {
                    MessageBox.Show("Не корректный ID коробки.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } catch {
                MessageBox.Show("Не корректный ID коробки.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            textBox9.Clear();
        }

        private void label7_Click(object sender, EventArgs e)
        {
            textBox9.Clear();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
            panel5.Visible = true;
            button6.Enabled = true;
            button5.Enabled = false;
            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            groupBox3.Enabled = false;
            textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            textBox13.Clear();
            numericUpDown1.Value = 1;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = true;
            panel5.Visible = false;
            button6.Enabled = false;
            button5.Enabled = true;
            groupBox1.Enabled = true;
            groupBox2.Enabled = true;
            groupBox3.Enabled = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы уверены что хотите удалить выбранную коробку со склада.", "Предупреждение.", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string grey_id = textBox9.Text;
                checkDelete(grey_id);
                //do something
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void label20_Click(object sender, EventArgs e)
        {
            textBox8.Clear();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            textBox14.Clear();
            textBox15.Clear();
            textBox16.Clear();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            textBox13.Clear();
            numericUpDown1.Value = 1;
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            //for (int x = 0; x < dataGridView1.ColumnCount; x++) // Проход по столбцам
            //{
            //    for (int y = 0; y < dataGridView1.RowCount; y++) // Проход по строкам
            //    {
            //        //if (Convert.ToInt32(dgvCal[x, y].Value) == 1) // Если в ячейке записано "1"
            //        //{
            //        using (Graphics g = e.Graphics)
            //        {
            //            g.DrawRectangle(Pens.Red, dataGridView1.GetCellDisplayRectangle(x, y, true)); // Нарисовать красную линию по краям ячейки, удовлетворяющей условию выше
            //        }
            //        //}
            //    }
            //}
        }

        private void dataGridView1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string grey_id = textBox8.Text;
            string boxNumber = "1";
            string partName = textBox10.Text;
            string partCode = textBox11.Text;
            string amount = numericUpDown1.Value.ToString();
            string net = textBox12.Text;
            string gross = textBox13.Text;
            string idOrder = textBox14.Text;
            string pl_id = textBox15.Text;
            string pl_date = "";
            string invoice_id = textBox16.Text;
            string invoice_date = "";
            int exists = BoxExists(grey_id);
            if (exists > 0) {
                addItem(boxNumber, partName, partCode, amount, net, gross, idOrder, grey_id, pl_id, pl_date, invoice_id, invoice_date);
                selectBoxItems(textBox8.Text);
            } else {
                MessageBox.Show("Коробки с указанным ID не существует.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
