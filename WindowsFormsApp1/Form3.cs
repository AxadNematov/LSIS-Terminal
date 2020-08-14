using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        string app_dir = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString());
        string app_dir_temp = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString()) + "\\Temp\\";

        private Form1 form1;

        public Form3(Form1 f1)
        {
            InitializeComponent();
            form1 = f1;
            InitGrid();
        }

        public void InitGrid()
        {
            dataGridView1.RowHeadersWidth = 35;
            //
            dataGridView1.ColumnCount = 8;
            //
            dataGridView1.Columns[0].HeaderText = "Название продукта";
            dataGridView1.Columns[1].HeaderText = "Код продукта";
            dataGridView1.Columns[2].HeaderText = "Кол-во";
            dataGridView1.Columns[3].HeaderText = "ID коробки";
            dataGridView1.Columns[4].HeaderText = "Секция";
            dataGridView1.Columns[5].HeaderText = "Ярус";
            dataGridView1.Columns[6].HeaderText = "Стелаж";
            dataGridView1.Columns[7].HeaderText = "Блок";
            //
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 100;
            dataGridView1.Columns[2].Width = 60;
            dataGridView1.Columns[3].Width = 90;
            dataGridView1.Columns[4].Width = 70;
            dataGridView1.Columns[5].Width = 70;
            dataGridView1.Columns[6].Width = 70;
            dataGridView1.Columns[7].Width = 70;
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        public class ItemList
        {
            public string boxNumber { get; set; }
            public string partName { get; set; }
            public string partCode { get; set; }
            public string amount { get; set; }
            public string blackId { get; set; }
            public string locationX { get; set; }
            public string locationY { get; set; }
            public string locationZ { get; set; }
            public string locationB { get; set; }

        }

        public void FindItems(string search, int s_type)
        {
            List<ItemList> items = new List<ItemList>();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            if (s_type == 0) {
                string connectionString = "SELECT items_in_boxes.box_numb, items_in_boxes.part_name, items_in_boxes.part_code, items_in_boxes.amount, items_in_boxes.black_id, Boxes.location_x, Boxes.location_y, Boxes.location_z, block FROM items_in_boxes JOIN Boxes ON items_in_boxes.black_id = Boxes.black_id WHERE LEN(items_in_boxes.black_id)>5 and items_in_boxes.part_code = @partcode";
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                comm.CommandText = connectionString;
                comm.Parameters.AddWithValue("@partcode", search);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    items.Add(new ItemList()
                    {
                        boxNumber = Convert.ToString(reader["box_numb"]),
                        partName = Convert.ToString(reader["part_name"]),
                        partCode = Convert.ToString(reader["part_code"]),
                        amount = Convert.ToString(reader["amount"]),
                        blackId = Convert.ToString(reader["black_id"]),
                        locationX = Convert.ToString(reader["location_x"]),
                        locationY = Convert.ToString(reader["location_y"]),
                        locationZ = Convert.ToString(reader["location_z"]),
                        locationB = Convert.ToString(reader["block"])
                    });
                }
            }
            //
            if (s_type == 1) {
                string connectionString = "SELECT items_in_boxes.box_numb, items_in_boxes.part_name, items_in_boxes.part_code, items_in_boxes.amount, items_in_boxes.black_id, Boxes.location_x, Boxes.location_y, Boxes.location_z, block FROM items_in_boxes JOIN Boxes ON items_in_boxes.black_id = Boxes.black_id WHERE LEN(items_in_boxes.black_id)>5 and items_in_boxes.part_name = @partname";
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                comm.CommandText = connectionString;
                comm.Parameters.AddWithValue("@partname", search);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    items.Add(new ItemList()
                    {
                        boxNumber = Convert.ToString(reader["box_numb"]),
                        partName = Convert.ToString(reader["part_name"]),
                        partCode = Convert.ToString(reader["part_code"]),
                        amount = Convert.ToString(reader["amount"]),
                        blackId = Convert.ToString(reader["black_id"]),
                        locationX = Convert.ToString(reader["location_x"]),
                        locationY = Convert.ToString(reader["location_y"]),
                        locationZ = Convert.ToString(reader["location_z"]),
                        locationB = Convert.ToString(reader["block"])
                    });
                }
            }
            connection.Close();
            FillDataGrid(items);
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

        public int MoveItem(string part_code, string grey_id_current, string grey_id_to, int qty_total, int qty_move)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            if (qty_move > qty_total || qty_move == 0) {
                MessageBox.Show("Недостаточное количество продуктов.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            //
            if (qty_move < qty_total && qty_move != qty_total) {
                List<string> list;
                // Decrease item qty from original box
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE items_in_boxes SET amount='" + (qty_total - qty_move) + "' WHERE part_code='" + part_code + "' and black_id='" + grey_id_current + "'";
                command.ExecuteNonQuery();
                // Get new box number
                int bn = -1;
                using (command = new SqlCommand("SELECT * FROM Boxes WHERE grey_id='" + grey_id_to + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            bn = Convert.ToInt32(reader.GetValue(1).ToString());
                        }
                    }
                }
                // Get all data for moving item
                list = new List<string>();
                using (command = new SqlCommand("SELECT * FROM items_in_boxes WHERE part_code='" + part_code + "' and black_id='" + grey_id_current + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(reader.GetValue(2).ToString());
                            list.Add(reader.GetValue(3).ToString());
                            list.Add(reader.GetValue(5).ToString());
                            list.Add(reader.GetValue(6).ToString());
                            list.Add(reader.GetValue(7).ToString());
                            list.Add(reader.GetValue(9).ToString());
                            list.Add(reader.GetValue(10).ToString());
                            list.Add(reader.GetValue(11).ToString());
                            list.Add(reader.GetValue(12).ToString());
                        }
                    }
                }
                //
                // Check if 'move to' box contains item
                int qty = 0;
                string p_code = "";
                string bl_id = "";
                using (command = new SqlCommand("SELECT black_id, amount, part_code FROM items_in_boxes WHERE black_id='" + grey_id_to + "' and part_code='" + part_code + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            bl_id = reader.GetValue(0).ToString();
                            qty = Convert.ToInt32(reader.GetValue(1).ToString());
                            p_code = reader.GetValue(2).ToString();
                        }
                    }
                }
                // Move to another box
                if (qty == 0)
                {
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO items_in_boxes(box_numb, part_name, part_code, amount, net, gross, id_order, black_id, pl_id, pl_date, invoice_id, invoice_date) VALUES(@box_numb, @part_name, @part_code, @amount, @net, @gross, @id_order, @black_id, @pl_id, @pl_date, @invoice_id, @invoice_date)";
                    command.Parameters.AddWithValue("@box_numb", bn);
                    command.Parameters.AddWithValue("@part_name", list[0]);
                    command.Parameters.AddWithValue("@part_code", list[1]);
                    command.Parameters.AddWithValue("@amount", qty_move);
                    command.Parameters.AddWithValue("@net", list[2]);
                    command.Parameters.AddWithValue("@gross", list[3]);
                    command.Parameters.AddWithValue("@id_order", list[4]);
                    command.Parameters.AddWithValue("@black_id", grey_id_to);
                    command.Parameters.AddWithValue("@pl_id", list[5]);
                    command.Parameters.AddWithValue("@pl_date", list[6]);
                    command.Parameters.AddWithValue("@invoice_id", list[7]);
                    command.Parameters.AddWithValue("@invoice_date", list[8]);
                    command.ExecuteNonQuery();
                }else
                {
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "UPDATE items_in_boxes SET amount='" + (qty + qty_move) + "' WHERE part_code='" + part_code + "' and black_id='" + grey_id_to + "'";
                    command.ExecuteNonQuery();
                }

            }
            //
            if (qty_move == qty_total)
            {
                List<string> list;
                // Decrease item qty from original box
                SqlCommand command = new SqlCommand();
                // Get new box number
                int bn = -1;
                using (command = new SqlCommand("SELECT * FROM Boxes WHERE grey_id='" + grey_id_to + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            bn = Convert.ToInt32(reader.GetValue(1).ToString());
                        }
                    }
                }
                // Get all data for moving item
                list = new List<string>();
                using (command = new SqlCommand("SELECT * FROM items_in_boxes WHERE part_code='" + part_code + "' and black_id='" + grey_id_current + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(reader.GetValue(2).ToString());
                            list.Add(reader.GetValue(3).ToString());
                            list.Add(reader.GetValue(5).ToString());
                            list.Add(reader.GetValue(6).ToString());
                            list.Add(reader.GetValue(7).ToString());
                            list.Add(reader.GetValue(9).ToString());
                            list.Add(reader.GetValue(10).ToString());
                            list.Add(reader.GetValue(11).ToString());
                            list.Add(reader.GetValue(12).ToString());
                        }
                    }
                }
                // Move to another box
                //command = new SqlCommand();
                //command.Connection = connection;
                //command.CommandType = CommandType.Text;
                //command.CommandText = "INSERT INTO items_in_boxes(box_numb, part_name, part_code, amount, net, gross, id_order, black_id, pl_id, pl_date, invoice_id, invoice_date) VALUES(@box_numb, @part_name, @part_code, @amount, @net, @gross, @id_order, @black_id, @pl_id, @pl_date, @invoice_id, @invoice_date)";
                //command.Parameters.AddWithValue("@box_numb", bn);
                //command.Parameters.AddWithValue("@part_name", list[0]);
                //command.Parameters.AddWithValue("@part_code", list[1]);
                //command.Parameters.AddWithValue("@amount", qty_move);
                //command.Parameters.AddWithValue("@net", list[2]);
                //command.Parameters.AddWithValue("@gross", list[3]);
                //command.Parameters.AddWithValue("@id_order", list[4]);
                //command.Parameters.AddWithValue("@black_id", grey_id_to);
                //command.Parameters.AddWithValue("@pl_id", list[5]);
                //command.Parameters.AddWithValue("@pl_date", list[6]);
                //command.Parameters.AddWithValue("@invoice_id", list[7]);
                //command.Parameters.AddWithValue("@invoice_date", list[8]);
                //command.ExecuteNonQuery();
                //
                int qty_in = 0;
                using (command = new SqlCommand("SELECT amount FROM items_in_boxes WHERE black_id='" + grey_id_to + "' and part_code='" + part_code + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            qty_in = Convert.ToInt32(reader.GetValue(0).ToString());
                        }
                    }
                }
                //
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE items_in_boxes SET pl_id='', pl_date='', invoice_id='', invoice_date='', amount='" + (qty_in + qty_move) + "' WHERE part_code='" + part_code + "' and black_id='" + grey_id_to + "'";
                command.ExecuteNonQuery();
                //
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "DELETE FROM items_in_boxes WHERE part_code='" + part_code + "' and black_id='" + grey_id_current + "'";
                command.ExecuteNonQuery();
            }
            //
            connection.Close();
            MessageBox.Show(qty_move + "  единиц продукта успешно перемещен из коробки  " + grey_id_current + "  в коробку  " + grey_id_to + ".", "Сообщение",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 1;
        }

        public void FillDataGrid(List<ItemList> list)
        {
            //dataGridView1.DataSource = list;
            dataGridView1.RowCount = list.Count;
            int total_amount = 0;
            for (int i = 0; i < list.Count; i++)
            {
                total_amount = total_amount + Convert.ToInt32(list[i].amount);
                dataGridView1.Rows[i].Cells[0].Value = list[i].partName;
                dataGridView1.Rows[i].Cells[1].Value = list[i].partCode;
                dataGridView1.Rows[i].Cells[2].Value = list[i].amount;
                dataGridView1.Rows[i].Cells[3].Value = list[i].blackId;
                dataGridView1.Rows[i].Cells[4].Value = list[i].locationX;
                dataGridView1.Rows[i].Cells[5].Value = list[i].locationY;
                dataGridView1.Rows[i].Cells[6].Value = list[i].locationZ;
                dataGridView1.Rows[i].Cells[7].Value = list[i].locationB;
                if (i % 2 == 0)
                {
                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[6].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[7].Style.BackColor = Color.WhiteSmoke;
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
                    dataGridView1.Rows[i].Cells[7].Style.BackColor = Color.White;
                }
            }
            if (list.Count != 0) {
                dataGridView1.RowCount = dataGridView1.RowCount + 1;
                dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1].Value = "Итого";
                dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[2].Value = total_amount;
            }
        }

        private void button8_search_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0) {
                //FindItems("62363172866", 0);
                FindItems(textBox1.Text, 0);
            }
            else {
                //FindItems("SUB ASS'Y,CTC,5A,MT,VL", 1);
                FindItems(textBox1.Text, 1);
            }
        }

        private void label19_Click_1(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void label4_Click(object sender, EventArgs e)
        {
            textBox5.Clear();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount != 0)
            {
                dataGridView1.Enabled = false;
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = false;
                groupBox2.Enabled = false;
                textBox2.Enabled = true;
                textBox5.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = true;
            groupBox2.Enabled = true;
            textBox2.Clear();
            textBox5.Clear();
            textBox2.Enabled = false;
            textBox5.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int qty;
            try {
                qty = Convert.ToInt32(textBox2.Text);
                //
                int selected = dataGridView1.CurrentCell.RowIndex;
                string part_code = dataGridView1.Rows[selected].Cells[1].Value.ToString();
                string grey_id_current = dataGridView1.Rows[selected].Cells[3].Value.ToString();
                string grey_id_to = textBox5.Text;
                string qty_total = dataGridView1.Rows[selected].Cells[2].Value.ToString();
                string qty_move = textBox2.Text;
                int exists = BoxExists(grey_id_to);
                if (grey_id_current == grey_id_to) {
                    MessageBox.Show("Невозможно переместить в коробку из этой же коробки.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (exists > 0) {
                    int ret = MoveItem(part_code, grey_id_current, grey_id_to, Convert.ToInt32(qty_total), Convert.ToInt32(qty_move));
                    if (ret == 0) { return; }
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button4.Enabled = true;
                    groupBox2.Enabled = true;
                    textBox2.Clear();
                    textBox5.Clear();
                    textBox2.Enabled = false;
                    textBox5.Enabled = false;
                    if (comboBox1.SelectedIndex == 0) {
                        FindItems(textBox1.Text, 0);
                    }else {
                        FindItems(textBox1.Text, 1);
                    }
                    dataGridView1.Enabled = true;
                }
                else {
                    MessageBox.Show("Коробки с указанным ID не существует.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            catch {
                MessageBox.Show("Указано не корректное количество.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
