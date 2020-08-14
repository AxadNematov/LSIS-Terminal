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
    public partial class Form5 : Form
    {
        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        private Form1 Frm1;
        int _move_box = 0;
        bool _product_scan = false;

        public Form5(Form1 f1)
        {
            InitializeComponent();
            Frm1 = f1;
            InitGrids();
        }

        public string TextBox_Item_Code
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }

        public string TextBox_Name
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }
        
        public string box_number { get; set; }
        public int quantity { get; set; }
        public int needed_on_contract { get; set; }
        
        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        public string MakePaletID()
        {
            // Генерация WHite ID
            string white_id;
            Random rnd = new Random();
            int valueFirst = rnd.Next(9999999, 99999999);
            white_id = valueFirst.ToString();
            return white_id;
        }

        public string GeneratePaletID(string ct_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            string pid = "";
            bool check_pid = false;
            while (check_pid == false)
            {
                pid = MakePaletID();
                SqlCommand check_pid_com = new SqlCommand("SELECT COUNT(*) FROM [collected_contract_items] WHERE ([no_poleta] = @pid)", connection);
                check_pid_com.Parameters.AddWithValue("@pid", pid);
                int idExist = (int)check_pid_com.ExecuteScalar();
                //
                if (idExist > 0)
                {
                    check_pid = true;
                }
                else
                {
                    check_pid = false;
                    break;
                }
            }
            return pid;
        }

        public void MoveItemWID(string box_from, string box_to, string item, int qty_move, int qty_in_box, string item_nm, string ct_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            //
            int qty_in_move_box = 0;
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT quantity FROM collected_contract_items WHERE item_code=@item and white_id=@box_to";
            comm.Parameters.AddWithValue("@item", item);
            comm.Parameters.AddWithValue("@box_to", box_to);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                qty_in_move_box = Convert.ToInt32(reader.GetValue(0).ToString());
            }
            reader.Close();
            //
            string palet_id = "";
            comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT no_poleta FROM collected_contract_items WHERE white_id=@box_to and no_poleta is not NULL";
            comm.Parameters.AddWithValue("@box_to", box_to);
            reader = comm.ExecuteReader();
            while (reader.Read())
            {
                palet_id = reader.GetValue(0).ToString();
            }
            reader.Close();
            if (qty_move < qty_in_box)
            {
                int left = qty_in_box - qty_move;
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE collected_contract_items SET quantity=@qty WHERE white_id=@box_from and item_code=@itemcode";
                command.Parameters.AddWithValue("@qty", left);
                command.Parameters.AddWithValue("@box_from", box_from);
                command.Parameters.AddWithValue("@itemcode", item);
                command.ExecuteNonQuery();
                //
                if (qty_in_move_box > 0)
                {
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "UPDATE collected_contract_items SET quantity=@qty WHERE white_id=@box_to and item_code=@itemcode";
                    command.Parameters.AddWithValue("@qty", qty_move + qty_in_move_box);
                    command.Parameters.AddWithValue("@box_to", box_to);
                    command.Parameters.AddWithValue("@itemcode", item);
                    command.ExecuteNonQuery();
                }
                else
                {
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO collected_contract_items (id_contract, white_id, item_code, name, quantity, no_poleta) VALUES(@contractid, @whiteid, @itemcode, @name, @quantity, @paleta)";
                    command.Parameters.AddWithValue("@contractid", ct_id);
                    command.Parameters.AddWithValue("@whiteid", box_to);
                    command.Parameters.AddWithValue("@itemcode", item);
                    command.Parameters.AddWithValue("@name", item_nm);
                    command.Parameters.AddWithValue("@quantity", qty_move);
                    command.Parameters.AddWithValue("@paleta", palet_id);
                    command.ExecuteNonQuery();
                }
            }
            if (qty_in_box == qty_move)
            {
                if (qty_in_move_box == 0)
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "UPDATE collected_contract_items SET white_id = @box_to WHERE white_id=@box_from and item_code=@itemcode";
                    command.Parameters.AddWithValue("@box_from", box_from);
                    command.Parameters.AddWithValue("@box_to", box_to);
                    command.Parameters.AddWithValue("@itemcode", item);
                    command.ExecuteNonQuery();
                }
                if (qty_in_move_box > 0)
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "DELETE FROM collected_contract_items WHERE white_id=@box_from and item_code=@itemcode";
                    command.Parameters.AddWithValue("@box_from", box_from);
                    command.Parameters.AddWithValue("@itemcode", item);
                    command.ExecuteNonQuery();
                    //
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "UPDATE collected_contract_items SET quantity=@qty WHERE white_id=@box_to and item_code=@itemcode";
                    command.Parameters.AddWithValue("@qty", qty_move + qty_in_move_box);
                    command.Parameters.AddWithValue("@box_to", box_to);
                    command.Parameters.AddWithValue("@itemcode", item);
                    command.ExecuteNonQuery();
                }
            }
            //    MessageBox.Show("Не достаточное количество продукта в коробке.", "Сообщение",
            //    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            GetProductsbyWID(box_from, ct_id, dataGridView3);
            GetProductsbyWID(box_to, ct_id, dataGridView6);
            connection.Close();
        }

        public void RemoveWIDfromPalet(string wid, string ct_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            //
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE [collected_contract_items] SET no_poleta=NULL WHERE id_contract='" + ct_id + "' and white_id='" + wid + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void AttachWIDtoPalet(string wid, string pid)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            //
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE [collected_contract_items] SET no_poleta='" + pid + "' WHERE white_id='" + wid + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void GetPaletBoxes(string palet_id, string ct_id, int type)
        {
            if (type == 0) {
                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                //
                List<string> list1 = new List<string>();
                List<string> list2 = new List<string>();
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                comm.CommandText = "SELECT DISTINCT white_id, id_contract FROM collected_contract_items WHERE no_poleta=@pid";
                comm.Parameters.AddWithValue("@pid", palet_id);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    list1.Add(reader.GetValue(0).ToString());
                    list2.Add(reader.GetValue(1).ToString());
                }
                reader.Close();
                //
                dataGridView10.RowCount = list1.Count;
                dataGridView10.RowHeadersWidth = 35;
                for (int n = 0; n < list1.Count; n++)
                {
                    dataGridView10.Rows[n].Cells[0].Value = n + 1;
                    dataGridView10.Rows[n].Cells[1].Value = list1[n];
                    dataGridView10.Rows[n].Cells[2].Value = list2[n];
                    if (n % 2 == 0)
                    {
                        dataGridView10.Rows[n].Cells[0].Style.BackColor = Color.WhiteSmoke;
                        dataGridView10.Rows[n].Cells[1].Style.BackColor = Color.WhiteSmoke;
                        dataGridView10.Rows[n].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    }
                    else
                    {
                        dataGridView10.Rows[n].Cells[0].Style.BackColor = Color.White;
                        dataGridView10.Rows[n].Cells[1].Style.BackColor = Color.White;
                        dataGridView10.Rows[n].Cells[2].Style.BackColor = Color.White;
                    }
                }
                connection.Close();
                //
            }
            else {
                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                //
                List<string> list1 = new List<string>();
                List<string> list2 = new List<string>();
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                //comm.CommandText = "SELECT DISTINCT white_id, id_contract FROM collected_contract_items WHERE no_poleta=@pid";
                comm.CommandText = "SELECT DISTINCT white_id, id_contract FROM collected_contract_items WHERE no_poleta=@pid and id_contract=@ct";
                comm.Parameters.AddWithValue("@pid", palet_id);
                comm.Parameters.AddWithValue("@ct", ct_id);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    list1.Add(reader.GetValue(0).ToString());
                    list2.Add(reader.GetValue(1).ToString());
                }
                reader.Close();
                //
                dataGridView10.RowCount = list1.Count;
                dataGridView10.RowHeadersWidth = 35;
                for (int n = 0; n < list1.Count; n++)
                {
                    dataGridView10.Rows[n].Cells[0].Value = n + 1;
                    dataGridView10.Rows[n].Cells[1].Value = list1[n];
                    dataGridView10.Rows[n].Cells[2].Value = list2[n];
                    if (n % 2 == 0)
                    {
                        dataGridView10.Rows[n].Cells[0].Style.BackColor = Color.WhiteSmoke;
                        dataGridView10.Rows[n].Cells[1].Style.BackColor = Color.WhiteSmoke;
                        dataGridView10.Rows[n].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    }
                    else
                    {
                        dataGridView10.Rows[n].Cells[0].Style.BackColor = Color.White;
                        dataGridView10.Rows[n].Cells[1].Style.BackColor = Color.White;
                        dataGridView10.Rows[n].Cells[2].Style.BackColor = Color.White;
                    }
                }
                connection.Close();
                //
            }
        }

        public void GetCTsList()
        {
            comboBox2.Items.Clear();
            comboBox2.Items.Insert(0, "( * )");
            comboBox3.Items.Clear();
            for (int i = 0; i < Frm1.dataGridView19.RowCount; i++) {
                comboBox2.Items.Insert(i + 1, Frm1.dataGridView19.Rows[i].Cells[1].Value.ToString());
                comboBox3.Items.Insert(i, Frm1.dataGridView19.Rows[i].Cells[1].Value.ToString());
            }
            timer_focus.Enabled = true;
        }

        public void GetPaletList(string ct_id)
        {
            try
            {
                string sql_str;
                string sql;
                if (ct_id == "")
                {
                    sql_str = "SELECT DISTINCT no_poleta FROM collected_contract_items WHERE (no_poleta is not NULL) AND (";
                    for (int i = 0; i < Frm1.dataGridView19.RowCount; i++)
                    {
                        ct_id = Frm1.dataGridView19.Rows[i].Cells[1].Value.ToString();
                        sql_str = sql_str + " id_contract='" + ct_id + "' OR";
                    }
                    sql = sql_str.Substring(0, sql_str.Length - 3);
                    sql = sql + ")";
                }
                else
                {
                    sql = "SELECT DISTINCT no_poleta FROM collected_contract_items WHERE id_contract='" + ct_id + "' AND no_poleta IS NOT NULL";
                }

                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                //
                List<string> list = new List<string>();
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                //comm.CommandText = "SELECT DISTINCT no_poleta FROM collected_contract_items WHERE id_contract=@ct_id and no_poleta!=null";
                comm.CommandText = sql;
                //comm.Parameters.AddWithValue("@ct_id", ct_id);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(reader.GetValue(0).ToString());
                }
                reader.Close();
                //
                dataGridView9.RowCount = list.Count;
                dataGridView9.RowHeadersWidth = 35;
                for (int n = 0; n < list.Count; n++)
                {
                    dataGridView9.Rows[n].Cells[0].Value = n + 1;
                    dataGridView9.Rows[n].Cells[1].Value = list[n];
                    if (n % 2 == 0)
                    {
                        dataGridView9.Rows[n].Cells[0].Style.BackColor = Color.WhiteSmoke;
                        dataGridView9.Rows[n].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    }
                    else
                    {
                        dataGridView9.Rows[n].Cells[0].Style.BackColor = Color.White;
                        dataGridView9.Rows[n].Cells[1].Style.BackColor = Color.White;
                    }
                }
                connection.Close();
                //
            }
            catch { }
        }

        public void GetBoxesbyCT(string ct_id, int count, DataGridView Grid)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            //
            List<string> list = new List<string>();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT DISTINCT white_id, id_contract FROM collected_contract_items WHERE id_contract=@ct_id AND no_poleta IS NULL";
            comm.Parameters.AddWithValue("@ct_id", ct_id);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                list.Add(reader.GetValue(0).ToString());
            }
            reader.Close();
            //
            Grid.RowCount = count + list.Count;
            Grid.RowHeadersWidth = 35;
            for (int n = 0; n < list.Count; n++)
            {
                Grid.Rows[n + count].Cells[0].Value = n + 1 + count;
                Grid.Rows[n + count].Cells[1].Value = list[n];
                Grid.Rows[n + count].Cells[2].Value = ct_id;
                if ((n + count) % 2 == 0)
                {
                    Grid.Rows[n + count].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    Grid.Rows[n + count].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    Grid.Rows[n + count].Cells[2].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    Grid.Rows[n + count].Cells[0].Style.BackColor = Color.White;
                    Grid.Rows[n + count].Cells[1].Style.BackColor = Color.White;
                    Grid.Rows[n + count].Cells[2].Style.BackColor = Color.White;
                }
            }
            //
            connection.Close();
        }

        public void GetWID_byCT(string ct_id)
        {
            List<string> list = new List<string>();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT DISTINCT white_id FROM collected_contract_items WHERE id_contract=@ct_id";
            comm.Parameters.AddWithValue("@ct_id", ct_id);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                list.Add(reader.GetValue(0).ToString());
            }
            reader.Close();
            //
            dataGridView4.RowCount = list.Count;
            dataGridView4.RowHeadersWidth = 35;
            for (int i = 0; i < list.Count; i++)
            {
                dataGridView4.Rows[i].Cells[0].Value = i + 1;
                dataGridView4.Rows[i].Cells[1].Value = list[i];
                if (i % 2 == 0)
                {
                    dataGridView4.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView4.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView4.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView4.Rows[i].Cells[1].Style.BackColor = Color.White;
                }
            }
        }

        public class Product
        {
            public string item_code { get; set; }
            public string name { get; set; }
            public string quantity { get; set; }
        }

        public void GetProductsbyWID(string wid, string ct_id, DataGridView Grid)
        {
            List<Product> list = new List<Product>();
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT item_code, name, quantity FROM collected_contract_items WHERE id_contract=@ct_id and white_id=@wid";
            comm.Parameters.AddWithValue("@ct_id", ct_id);
            comm.Parameters.AddWithValue("@wid", wid);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                Product item = new Product();
                item.item_code = reader.GetValue(0).ToString();
                item.name = reader.GetValue(1).ToString();
                item.quantity = reader.GetValue(2).ToString();
                list.Add(item);
            }
            reader.Close();
            //
            Grid.RowCount = list.Count;
            Grid.RowHeadersWidth = 35;
            for (int i = 0; i < list.Count; i++)
            {
                Grid.Rows[i].Cells[0].Value = i + 1;
                Grid.Rows[i].Cells[1].Value = list[i].item_code;
                Grid.Rows[i].Cells[2].Value = list[i].name;
                Grid.Rows[i].Cells[3].Value = list[i].quantity;
                if (i % 2 == 0)
                {
                    Grid.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    Grid.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    Grid.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    Grid.Rows[i].Cells[3].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    Grid.Rows[i].Cells[0].Style.BackColor = Color.White;
                    Grid.Rows[i].Cells[1].Style.BackColor = Color.White;
                    Grid.Rows[i].Cells[2].Style.BackColor = Color.White;
                    Grid.Rows[i].Cells[3].Style.BackColor = Color.White;
                }
            }
        }

        public void MoveItemtoGrid(DataGridView Grid_from, DataGridView Grid_to)
        {
            int selected = 0;
            try
            {
                selected = Grid_from.CurrentCell.RowIndex;
                string wid = Grid_from.Rows[selected].Cells[1].Value.ToString();
                Grid_from.Rows.RemoveAt(selected);
                //
                Grid_to.RowCount = Grid_to.RowCount + 1;
                Grid_to.Rows[Grid_to.RowCount - 1].Cells[0].Value = Grid_to.RowCount;
                Grid_to.Rows[Grid_to.RowCount - 1].Cells[1].Value = wid;
                if (Grid_to.RowCount % 2 == 0)
                {
                    Grid_to.Rows[Grid_to.RowCount - 1].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    Grid_to.Rows[Grid_to.RowCount - 1].Cells[1].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    Grid_to.Rows[Grid_to.RowCount - 1].Cells[0].Style.BackColor = Color.White;
                    Grid_to.Rows[Grid_to.RowCount - 1].Cells[1].Style.BackColor = Color.White;
                }
            }
            catch { }
        }

        public void InitGrids()
        {
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 3;

            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "Id Контракта";
            dataGridView1.Columns[2].HeaderText = "Заказчик";

            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 120;
            //
            //
            dataGridView4.RowCount = 0;
            dataGridView4.ColumnCount = 2;

            dataGridView4.Columns[0].HeaderText = "#";
            dataGridView4.Columns[1].HeaderText = "ID коробки";

            dataGridView4.Columns[0].Width = 40;
            dataGridView4.Columns[1].Width = 90;
            //
            //
            dataGridView2.RowCount = 0;
            dataGridView2.ColumnCount = 3;

            dataGridView2.Columns[0].HeaderText = "#";
            dataGridView2.Columns[1].HeaderText = "ID Коробки";
            dataGridView2.Columns[2].HeaderText = "ID Контракта";

            dataGridView2.Columns[0].Width = 40;
            dataGridView2.Columns[1].Width = 90;
            dataGridView2.Columns[2].Width = 120;
            //
            //
            dataGridView3.RowCount = 0;
            dataGridView3.ColumnCount = 4;

            dataGridView3.Columns[0].HeaderText = "#";
            dataGridView3.Columns[1].HeaderText = "Код продукта";
            dataGridView3.Columns[2].HeaderText = "Название продукта";
            dataGridView3.Columns[3].HeaderText = "Кол-во";

            dataGridView3.Columns[0].Width = 40;
            dataGridView3.Columns[1].Width = 100;
            dataGridView3.Columns[2].Width = 220;
            dataGridView3.Columns[3].Width = 70;
            //
            //
            dataGridView7.RowCount = 0;
            dataGridView7.ColumnCount = 3;

            dataGridView7.Columns[0].HeaderText = "#";
            dataGridView7.Columns[1].HeaderText = "ID Коробки";
            dataGridView7.Columns[2].HeaderText = "ID Контракта";

            dataGridView7.Columns[0].Width = 40;
            dataGridView7.Columns[1].Width = 90;
            dataGridView7.Columns[2].Width = 120;
            //
            //
            dataGridView8.RowCount = 0;
            dataGridView8.ColumnCount = 3;

            dataGridView8.Columns[0].HeaderText = "#";
            dataGridView8.Columns[1].HeaderText = "ID Коробки";
            dataGridView8.Columns[2].HeaderText = "ID Контракта";

            dataGridView8.Columns[0].Width = 40;
            dataGridView8.Columns[1].Width = 90;
            dataGridView8.Columns[2].Width = 120;
            dataGridView8.Columns[2].Visible = false;

            //
            //
            dataGridView6.RowCount = 0;
            dataGridView6.ColumnCount = 4;

            dataGridView6.Columns[0].HeaderText = "#";
            dataGridView6.Columns[1].HeaderText = "Код продукта";
            dataGridView6.Columns[2].HeaderText = "Название продукта";
            dataGridView6.Columns[3].HeaderText = "Кол-во";

            dataGridView6.Columns[0].Width = 40;
            dataGridView6.Columns[1].Width = 100;
            dataGridView6.Columns[2].Width = 220;
            dataGridView6.Columns[3].Width = 70;
            //
            //
            dataGridView5.RowCount = 0;
            dataGridView5.ColumnCount = 4;

            dataGridView5.Columns[0].HeaderText = "#";
            dataGridView5.Columns[1].HeaderText = "Код продукта";
            dataGridView5.Columns[2].HeaderText = "Название продукта";
            dataGridView5.Columns[3].HeaderText = "Кол-во";

            dataGridView5.Columns[0].Width = 40;
            dataGridView5.Columns[1].Width = 100;
            dataGridView5.Columns[2].Width = 250;
            dataGridView5.Columns[3].Width = 70;
            //
            //
            dataGridView9.RowCount = 0;
            dataGridView9.ColumnCount = 2;

            dataGridView9.Columns[0].HeaderText = "#";
            dataGridView9.Columns[1].HeaderText = "ID Палеты";

            dataGridView9.Columns[0].Width = 40;
            dataGridView9.Columns[1].Width = 90;
            //
            //
            dataGridView10.RowCount = 0;
            dataGridView10.ColumnCount = 3;

            dataGridView10.Columns[0].HeaderText = "#";
            dataGridView10.Columns[1].HeaderText = "ID Коробки";
            dataGridView10.Columns[2].HeaderText = "ID Контракта";

            dataGridView10.Columns[0].Width = 40;
            dataGridView10.Columns[1].Width = 90;
            dataGridView10.Columns[2].Width = 110;
        }

        public int CheckBoxCT(string ct_id, string wid)
        {
            int result;
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "SELECT COUNT(*) FROM collected_contract_items WHERE white_id=@wid and id_contract!=@ct_id";
            command1.Parameters.AddWithValue("@wid", wid);
            command1.Parameters.AddWithValue("@ct_id", ct_id);
            result = Convert.ToInt32(command1.ExecuteScalar());
            connection.Close();
            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (_product_scan == true) {
                collectItem();
            //}else{
            //    MessageBox.Show("Не верный ID продукта.", "Сообщение",
            //    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
        }

        public void collectItem()
        {
            string contractId = "";
            int selected = 0;
            try
            {
                selected = dataGridView1.CurrentCell.RowIndex;
                contractId = dataGridView1.Rows[selected].Cells[1].Value.ToString();
            }
            catch { }
            //
            for(int n=0; n<Frm1.dataGridView16.RowCount; n++)
            {
                string contract = "";
                string item_code = "";
                try { contract = Frm1.dataGridView16.Rows[n].Cells[1].Value.ToString(); } catch { contract = ""; }
                try { item_code = Frm1.dataGridView16.Rows[n].Cells[2].Value.ToString(); } catch { item_code = ""; }
                if(contract==contractId && item_code==TextBox_Item_Code)
                {
                    needed_on_contract = Convert.ToInt32(Frm1.dataGridView16.Rows[n].Cells[5].Value.ToString());
                }
            }

            int qty = 0;
            try { qty = Convert.ToInt32(numericUpDown1.Value); } catch { qty = 0; }
            if (qty > quantity)
            {
                MessageBox.Show("ib box has only "+quantity.ToString()+" goods", "???",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if(qty>needed_on_contract)
            {
                MessageBox.Show("we need only " + needed_on_contract.ToString() + " goods", "???",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                int existingQty = 0;
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                comm.CommandText = "SELECT quantity FROM collected_contract_items WHERE white_id=@whiteid and item_code=@itemcode";
                comm.Parameters.AddWithValue("@whiteid", textBox4.Text);
                comm.Parameters.AddWithValue("@itemcode", TextBox_Item_Code);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    existingQty = Convert.ToInt32(reader.GetValue(0));
                }
                reader.Close();

                if (existingQty > 0)
                {
                    existingQty = existingQty + qty;
                    SqlCommand command3 = new SqlCommand();
                    command3.Connection = connection;
                    command3.CommandType = CommandType.Text;
                    command3.CommandText = "update collected_contract_items set quantity = @amount where white_id=@whiteid and item_code=@itemcode";
                    command3.Parameters.AddWithValue("@amount", existingQty);
                    command3.Parameters.AddWithValue("@whiteid", textBox4.Text);
                    command3.Parameters.AddWithValue("@itemcode", TextBox_Item_Code);
                    command3.ExecuteNonQuery();
                }
                else
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO collected_contract_items(id_contract, white_id, item_code, name, quantity) VALUES(@contractid, @whiteid, @itemcode, @name, @quantity)";
                    command.Parameters.AddWithValue("@contractid", contractId);
                    command.Parameters.AddWithValue("@whiteid", textBox4.Text);
                    command.Parameters.AddWithValue("@itemcode", TextBox_Item_Code);
                    command.Parameters.AddWithValue("@name", TextBox_Name);
                    command.Parameters.AddWithValue("@quantity", qty);
                    command.ExecuteNonQuery();
                }

                int temp_qty = 0;
                temp_qty = quantity - qty;

                SqlCommand command2 = new SqlCommand();
                command2.Connection = connection;
                command2.CommandType = CommandType.Text;
                command2.CommandText = "update items_in_boxes set amount = @amount where black_id=@blackid and part_code=@partcode";
                command2.Parameters.AddWithValue("@amount", temp_qty);
                command2.Parameters.AddWithValue("@blackid", box_number);
                command2.Parameters.AddWithValue("@partcode", TextBox_Item_Code);
                command2.ExecuteNonQuery();

                connection.Close();
                //string str = "processing: \nfrom box: " + box_number + "\nto box: " + textBox4.Text + "\nitem code:" + TextBox_Item_Code + "\nname: " + TextBox_Name + "\nquantity: " + textBox3.Text + "\ncontract: " + contractId;
                //MessageBox.Show(str, "",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
                for (int i = 0; i < Frm1.dataGridView16.RowCount; i++)
                {
                    string id_contract = "";
                    string item_code = "";
                    try { id_contract = Frm1.dataGridView16.Rows[i].Cells[1].Value.ToString(); } catch { id_contract = ""; }
                    try { item_code = Frm1.dataGridView16.Rows[i].Cells[2].Value.ToString(); } catch { item_code = ""; }
                    if (id_contract == contractId && item_code == TextBox_Item_Code)
                    {
                        Frm1.dataGridView16.Rows[i].Cells[5].Value = Convert.ToInt32(Frm1.dataGridView16.Rows[i].Cells[5].Value) - qty;
                    }
                    if (Convert.ToInt32(Frm1.dataGridView16.Rows[i].Cells[5].Value) == 0 && id_contract.Length>0)
                    {
                        Frm1.dataGridView16.Rows[i].Cells[6].Style.BackColor = Color.Green;
                    }
                }
            }
            Frm1.findBoxes();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Focus();
            dataGridView5.RowCount = 0;
            int selected = 0;
            try
            {
                selected = dataGridView1.CurrentCell.RowIndex;
                string ct_id = dataGridView1.Rows[selected].Cells[1].Value.ToString();
                GetWID_byCT(ct_id);
                for (int n = 0; n < Frm1.dataGridView16.RowCount; n++)
                {
                    string contract = "";
                    string item_code = "";
                    try { contract = Frm1.dataGridView16.Rows[n].Cells[1].Value.ToString(); } catch { contract = ""; }
                    try { item_code = Frm1.dataGridView16.Rows[n].Cells[2].Value.ToString(); } catch { item_code = ""; }
                    if (contract == ct_id && item_code == TextBox_Item_Code)
                    {
                        for(int k=0; k<Frm1.dataGridView16.RowCount; k++)
                        {
                            string check_contract_id = "";
                            try { check_contract_id = Frm1.dataGridView16.Rows[k].Cells[1].Value.ToString(); } catch { check_contract_id = ""; }
                            if (check_contract_id.Length>0)
                            {
                                if (k % 2 == 0)
                                {
                                    Frm1.dataGridView16.Rows[k].Cells[0].Style.BackColor = Color.WhiteSmoke;
                                    Frm1.dataGridView16.Rows[k].Cells[1].Style.BackColor = Color.WhiteSmoke;
                                    Frm1.dataGridView16.Rows[k].Cells[2].Style.BackColor = Color.WhiteSmoke;
                                    Frm1.dataGridView16.Rows[k].Cells[3].Style.BackColor = Color.WhiteSmoke;
                                    Frm1.dataGridView16.Rows[k].Cells[4].Style.BackColor = Color.WhiteSmoke;
                                    Frm1.dataGridView16.Rows[k].Cells[5].Style.BackColor = Color.WhiteSmoke;
                                }
                                else
                                {
                                    Frm1.dataGridView16.Rows[k].Cells[0].Style.BackColor = Color.White;
                                    Frm1.dataGridView16.Rows[k].Cells[1].Style.BackColor = Color.White;
                                    Frm1.dataGridView16.Rows[k].Cells[2].Style.BackColor = Color.White;
                                    Frm1.dataGridView16.Rows[k].Cells[3].Style.BackColor = Color.White;
                                    Frm1.dataGridView16.Rows[k].Cells[4].Style.BackColor = Color.White;
                                    Frm1.dataGridView16.Rows[k].Cells[5].Style.BackColor = Color.White;
                                }
                            }
                        }
                        Frm1.dataGridView16.Rows[n].Cells[0].Style.BackColor = Color.Bisque;
                        Frm1.dataGridView16.Rows[n].Cells[1].Style.BackColor = Color.Bisque;
                        Frm1.dataGridView16.Rows[n].Cells[2].Style.BackColor = Color.Bisque;
                        Frm1.dataGridView16.Rows[n].Cells[3].Style.BackColor = Color.Bisque;
                        Frm1.dataGridView16.Rows[n].Cells[4].Style.BackColor = Color.Bisque;
                        Frm1.dataGridView16.Rows[n].Cells[5].Style.BackColor = Color.Bisque;
                    }
                }
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
            textBox4.Clear();
        }

        private void Form5_Activated(object sender, EventArgs e)
        {
            textBox6.Focus();
            GetCTsList();
            _product_scan = false;
            pictureBox1.Visible = false;
            pictureBox2.Visible = true;
            numericUpDown1.Value = 0;
            //MessageBox.Show("Focused Form5/", "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Information);
            textBox6.Focus();
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            ////dataGridView2.RowCount = 0;
            //string ct = comboBox2.Text;
            //if (ct != "( * )")
            //{
            //    GetBoxesbyCT(ct, 0, dataGridView2);
            //    GetBoxesbyCT(ct, 0, dataGridView7);
            //}
            //else
            //{
            //    for (int i = 0; i < comboBox2.Items.Count; i++)
            //    {
            //        ct = comboBox2.Items[i].ToString();
            //        GetBoxesbyCT(ct, dataGridView2.RowCount, dataGridView2);
            //        GetBoxesbyCT(ct, dataGridView7.RowCount, dataGridView7);
            //    }
            //}
            //timer_focus.Enabled = true;
        }

        private void label7_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
        }

        private void label7_Click_1(object sender, EventArgs e)
        {
            textBox4.Clear();
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Focus();
            int selected = 0;
            try
            {
                selected = dataGridView1.CurrentCell.RowIndex;
                string ct_id = dataGridView1.Rows[selected].Cells[1].Value.ToString();
                selected = dataGridView4.CurrentCell.RowIndex;
                string white_id = dataGridView4.Rows[selected].Cells[1].Value.ToString();
                textBox4.Text = white_id;
                GetProductsbyWID(white_id, ct_id, dataGridView5);
                //MessageBox.Show("", "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            textBox7.Focus();
            int selected = 0;
            try
            {
                selected = dataGridView2.CurrentCell.RowIndex;
                string white_id = dataGridView2.Rows[selected].Cells[1].Value.ToString();
                string ct_id = dataGridView2.Rows[selected].Cells[2].Value.ToString();
                GetProductsbyWID(white_id, ct_id, dataGridView3);
                if (_move_box == 1) {
                    textBox3.Text = white_id;
                }
                else
                {
                    //textBox5.Text = white_id;
                }
                //MessageBox.Show("", "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "( * )" && comboBox2.Text != "") {
                textBox3.Clear();
                textBox5.Clear();
                textBox8.Clear();
                button2.Enabled = false;
                button3.Enabled = true;
                _move_box = 1;
                textBox7.Enabled = true;
                comboBox2.Enabled = false;
            }
            else {
                MessageBox.Show("Необходимо отфильтровать по одному контракту.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            textBox7.Focus();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBox2.Enabled = true;
            numericUpDown2.Value = 0;
            textBox3.Clear();
            textBox5.Clear();
            textBox8.Clear();
            button2.Enabled = true;
            button3.Enabled = false;
            button4.Enabled = false;
            button8.Enabled = false;
            _move_box = 0;
            dataGridView3.Enabled = true;
            dataGridView6.Enabled = true;
            dataGridView7.Enabled = true;
            textBox7.Focus();
            textBox7.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox5.Clear();
            button2.Enabled = true;
            button3.Enabled = false;
            button4.Enabled = false;
            _move_box = 0;
            dataGridView3.Enabled = true;
            textBox7.Focus();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            timer1.Enabled = true;
        }

         private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            string scan = textBox6.Text;
            textBox6.Clear();
            int selected;
            int check;
            //MessageBox.Show(textBox6.Text, "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Warning);
            try
            {
                if (scan.IndexOf("LSDZ") != -1)
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        string ct_id = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        if (ct_id == scan)
                        {
                            dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
                    }
                }
                if (scan.Length == 5)
                {
                    for (int i = 0; i < dataGridView4.RowCount; i++)
                    {
                        string white_id = dataGridView4.Rows[i].Cells[1].Value.ToString();
                        if (white_id == scan)
                        {
                            dataGridView4.CurrentCell = dataGridView4.Rows[i].Cells[0];
                            dataGridView4.Rows[i].Selected = true;
                            break;
                        }
                    }
                    //
                    selected = dataGridView1.CurrentCell.RowIndex;
                    string ct_id = dataGridView1.Rows[selected].Cells[1].Value.ToString();
                    check = CheckBoxCT(ct_id, scan);
                    if (check == 0)
                    {
                        textBox4.Text = scan;
                    }
                    else
                    {
                        MessageBox.Show("Указанная коробка привязанна к другому контракту.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                if (scan.Length == 10)
                {
                    if (dataGridView1.RowCount != 0)
                    {
                        if (textBox4.Text != "")
                        {
                            if (textBox1.Text == scan)
                            {
                                pictureBox1.Visible = true;
                                pictureBox2.Visible = false;
                                _product_scan = true;
                                try
                                {
                                    int a = Convert.ToInt32(numericUpDown1.Value);
                                    a = a + 1;
                                    numericUpDown1.Value = a;
                                }
                                catch { }
                            }
                            else
                            {
                                pictureBox1.Visible = false;
                                pictureBox2.Visible = true;
                                _product_scan = false;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Необходимо выбрать коробку.", "Сообщение",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Необходимо выбрать контракт.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch { }
            timer1.Enabled = false;
            //
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(numericUpDown1.Value);
            try{
                a = a + 1;
                numericUpDown1.Value = a;
            }
            catch { }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(numericUpDown1.Value);
            try{
                a = a - 1;
                numericUpDown1.Value = a;
            }
            catch { }
        }

        private void button5_MouseDown(object sender, MouseEventArgs e)
        {
            timer_plus.Enabled = true;
        }

        private void button5_MouseUp(object sender, MouseEventArgs e)
        {
            timer_plus.Enabled = false;
            textBox6.Focus();
        }

        private void timer_plus_Tick(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                int a = Convert.ToInt32(numericUpDown1.Value);
                try
                {
                    a = a + 1;
                    numericUpDown1.Value = a;
                }
                catch { }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                int a = Convert.ToInt32(numericUpDown2.Value);
                try
                {
                    a = a + 1;
                    numericUpDown2.Value = a;
                }
                catch { }
            }
        }

        private void button6_MouseDown(object sender, MouseEventArgs e)
        {
            timer_minus.Enabled = true;
        }

        private void button6_MouseUp(object sender, MouseEventArgs e)
        {
            timer_minus.Enabled = false;
            textBox6.Focus();
        }

        private void timer_minus_Tick(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                int a = Convert.ToInt32(numericUpDown1.Value);
                try
                {
                    a = a - 1;
                    numericUpDown1.Value = a;
                }
                catch { }
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                int a = Convert.ToInt32(numericUpDown2.Value);
                try
                {
                    a = a - 1;
                    numericUpDown2.Value = a;
                }
                catch { }
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void numericUpDown1_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void button1_MouseUp(object sender, MouseEventArgs e)
        {
            textBox6.Focus();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            numericUpDown1.Value = 0;
        }

        private void button7_MouseUp(object sender, MouseEventArgs e)
        {
            textBox6.Focus();
        }

        private void timer_focus_Tick(object sender, EventArgs e)
        {
            timer_focus.Enabled = false;
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                textBox6.Focus();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                textBox7.Focus();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
            {
                if (dataGridView9.RowCount == 0) {
                    GetPaletList("");
                }
                textBox10.Focus();
            }
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            //{
            timer_focus.Enabled = true;
            //}
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            //{
            //    timer_focus.Enabled = true;
            //}
        }

        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            textBox7.Focus();
            int selected = 0;
            try
            {
                selected = dataGridView7.CurrentCell.RowIndex;
                string white_id = dataGridView7.Rows[selected].Cells[1].Value.ToString();
                string ct_id = dataGridView7.Rows[selected].Cells[2].Value.ToString();
                GetProductsbyWID(white_id, ct_id, dataGridView6);
                if (_move_box == 1)
                {
                    textBox5.Text = white_id;
                }
                else
                {
                    //textBox5.Text = white_id;
                }
                //MessageBox.Show("", "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(numericUpDown2.Value);
            try
            {
                a = a + 1;
                numericUpDown2.Value = a;
            }
            catch { }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(numericUpDown2.Value);
            try
            {
                a = a - 1;
                numericUpDown2.Value = a;
            }
            catch { }
        }

        private void button11_MouseDown(object sender, MouseEventArgs e)
        {
            timer_plus.Enabled = true;
        }

        private void button11_MouseUp(object sender, MouseEventArgs e)
        {
            timer_plus.Enabled = false;
            textBox7.Focus();
        }

        private void button10_MouseDown(object sender, MouseEventArgs e)
        {
            timer_minus.Enabled = true;
        }

        private void button10_MouseUp(object sender, MouseEventArgs e)
        {
            timer_minus.Enabled = false;
            textBox7.Focus();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            numericUpDown2.Value = 0;
        }

        private void button9_MouseUp(object sender, MouseEventArgs e)
        {
            textBox7.Focus();
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox7.Focus();
            int selected = 0;
            try
            {
                selected = dataGridView3.CurrentCell.RowIndex;
                string item_code = dataGridView3.Rows[selected].Cells[1].Value.ToString();
                if (_move_box == 1)
                {
                    textBox8.Text = item_code;
                    int a = Convert.ToInt32(numericUpDown2.Value);
                    try
                    {
                        a = a = 1;
                        numericUpDown2.Value = a;
                    }
                    catch { }
                }
                //MessageBox.Show("", "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            textBox7.Focus();
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            //textBox7.Focus();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int selected = 0;
            //
            string box_from = "";
            string box_to = "";
            string item = "";
            string item_nm = "";
            string ct_id = "";
            int qty_move = 0;
            int qty_in_box = 0;
            try
            {
                selected = dataGridView2.CurrentCell.RowIndex;
                box_from = dataGridView2.Rows[selected].Cells[1].Value.ToString();
                //
                selected = dataGridView7.CurrentCell.RowIndex;
                box_to = dataGridView7.Rows[selected].Cells[1].Value.ToString();
                //
                selected = dataGridView3.CurrentCell.RowIndex;
                item = dataGridView3.Rows[selected].Cells[1].Value.ToString();
                //
                selected = dataGridView3.CurrentCell.RowIndex;
                item_nm = dataGridView3.Rows[selected].Cells[2].Value.ToString();
                //
                selected = dataGridView3.CurrentCell.RowIndex;
                qty_in_box = Convert.ToInt32(dataGridView3.Rows[selected].Cells[3].Value.ToString());
                //
                selected = dataGridView2.CurrentCell.RowIndex;
                ct_id = dataGridView2.Rows[selected].Cells[2].Value.ToString();
                //
                qty_move = Convert.ToInt32(numericUpDown2.Value.ToString());
                //
                MoveItemWID(box_from, box_to, item, qty_move, qty_in_box, item_nm, ct_id);
                //
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
            comboBox2.Enabled = true;
            numericUpDown2.Value = 0;
            textBox3.Clear();
            textBox5.Clear();
            textBox8.Clear();
            button2.Enabled = true;
            button3.Enabled = false;
            button8.Enabled = false;
            _move_box = 0;
            dataGridView3.Enabled = true;
            dataGridView6.Enabled = true;
            dataGridView7.Enabled = true;
            textBox7.Focus();
            textBox7.Enabled = false;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (_move_box == 1){
                timer2.Enabled = true;
            }
            else {
                textBox7.Clear();
                MessageBox.Show("Для перемещения продукта из коробки в коробку необходимо запустить режим перемещения.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            timer_focus.Enabled = true;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            string scan = textBox7.Text;
            textBox7.Clear();
            int selected;
            int check;
            //MessageBox.Show(textBox6.Text, "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Warning);
            try
            {
                if (scan.Length == 5)
                {
                    if (textBox3.Text == "") {
                        for (int i = 0; i < dataGridView2.RowCount; i++)
                        {
                            string white_id = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            if (white_id == scan)
                            {
                                dataGridView2.CurrentCell = dataGridView2.Rows[i].Cells[0];
                                dataGridView2.Rows[i].Selected = true;
                                textBox3.Text = scan;
                                break;
                            }
                        }
                    }
                    if (textBox3.Text != "" && textBox8.Text != "")
                    {
                        for (int i = 0; i < dataGridView7.RowCount; i++)
                        {
                            string white_id = dataGridView7.Rows[i].Cells[1].Value.ToString();
                            if (white_id == scan)
                            {
                                dataGridView7.CurrentCell = dataGridView7.Rows[i].Cells[0];
                                dataGridView7.Rows[i].Selected = true;
                                textBox5.Text = scan;
                                dataGridView6.Enabled = true;
                                dataGridView7.Enabled = true;
                                break;
                            }
                        }
                    }
                }
                if (scan.Length == 10)
                {
                    if (textBox3.Text != "")
                    {
                        for (int i = 0; i < dataGridView3.RowCount; i++)
                        {
                            string item_code = dataGridView3.Rows[i].Cells[1].Value.ToString();
                            if (item_code == scan)
                            {
                                dataGridView3.CurrentCell = dataGridView3.Rows[i].Cells[0];
                                dataGridView3.Rows[i].Selected = true;
                                textBox8.Text = scan;
                                break;
                            }
                        }

                        if (textBox8.Text != "")
                        {
                            try
                            {
                                int a = Convert.ToInt32(numericUpDown2.Value);
                                a = a + 1;
                                numericUpDown2.Value = a;
                            }
                            catch { }
                        }
                    }else{
                        MessageBox.Show("Сначала выберите коробку из которой необходимо переместить продукт.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch { }
            timer2.Enabled = false;
            //
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown2.Value == 0) {
                button4.Enabled = false;
                button8.Enabled = false;
                dataGridView6.Enabled = false;
                dataGridView7.Enabled = false;
                textBox5.Clear();
            }
            else
            {
                button4.Enabled = true;
                button8.Enabled = true;
                dataGridView6.Enabled = true;
                dataGridView7.Enabled = true;
            }
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void dataGridView4_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void dataGridView5_Click(object sender, EventArgs e)
        {
            textBox6.Focus();
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            textBox7.Focus();
        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {
            textBox7.Focus();
        }

        private void dataGridView7_Click(object sender, EventArgs e)
        {
            textBox7.Focus();
        }

        private void dataGridView6_Click(object sender, EventArgs e)
        {
            textBox7.Focus();
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            timer_focus.Enabled = true;
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            string ct = comboBox3.Text;
            label12.Text = ct;
            if (ct != "( * )")
            {
                GetBoxesbyCT(ct, 0, dataGridView8);
                //GetPaletList(ct);
            }
            //else
            //{
            //    for (int i = 0; i < comboBox3.Items.Count; i++)
            //    {
            //        ct = comboBox3.Items[i].ToString();
            //        GetBoxesbyCT(ct, dataGridView8.RowCount, dataGridView8);
            //    }
            //}
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            //dataGridView2.RowCount = 0;
            string ct = comboBox2.Text;
            if (ct != "( * )")
            {
                GetBoxesbyCT(ct, 0, dataGridView2);
                GetBoxesbyCT(ct, 0, dataGridView7);
            }
            else
            {
                for (int i = 0; i < comboBox2.Items.Count; i++)
                {
                    ct = comboBox2.Items[i].ToString();
                    GetBoxesbyCT(ct, dataGridView2.RowCount, dataGridView2);
                    GetBoxesbyCT(ct, dataGridView7.RowCount, dataGridView7);
                }
            }
            timer_focus.Enabled = true;
        }

        private void dataGridView8_Click(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void dataGridView9_Click(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void dataGridView10_Click(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView8.CurrentCell.RowIndex;
                string wid = dataGridView8.Rows[selected].Cells[1].Value.ToString();
                selected = dataGridView9.CurrentCell.RowIndex;
                string pid = dataGridView9.Rows[selected].Cells[1].Value.ToString();
                //
                AttachWIDtoPalet(wid, pid);
                GetBoxesbyCT(label12.Text, 0, dataGridView8);
                //
                GetPaletBoxes(pid, label12.Text, 1);
                //
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
            comboBox3.Text = label12.Text;
            textBox10.Focus();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView10.CurrentCell.RowIndex;
                string wid = dataGridView10.Rows[selected].Cells[1].Value.ToString();
                string ct_id = dataGridView10.Rows[selected].Cells[2].Value.ToString();
                //
                RemoveWIDfromPalet(wid, ct_id);
                GetBoxesbyCT(ct_id, 0, dataGridView8);
                //
                selected = dataGridView9.CurrentCell.RowIndex;
                string pid = dataGridView9.Rows[selected].Cells[1].Value.ToString();
                //
                GetPaletBoxes(pid, label12.Text, 1);
                //
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
            comboBox3.Text = label12.Text;
            textBox10.Focus();
        }

        private void dataGridView11_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            comboBox3.Enabled = false;
        }

        private void dataGridView11_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {

        }

        private void dataGridView11_Click(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false;
            string scan = textBox10.Text;
            textBox10.Clear();
            int selected;
            int check;
            //MessageBox.Show(textBox6.Text, "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Warning);
            try
            {
                //
                if (scan.Length == 8)
                {
                    bool log = true;
                    for (int i = 0; i < dataGridView9.RowCount; i++)
                    {
                        string red_id = dataGridView9.Rows[i].Cells[1].Value.ToString();
                        if (red_id == scan)
                        {
                            textBox11.Text = scan;
                            dataGridView9.CurrentCell = dataGridView9.Rows[i].Cells[0];
                            dataGridView9.Rows[i].Selected = true;
                            log = false;
                            break;
                        }
                    }
                    if (log == true)
                    {
                        MessageBox.Show("Указанной палеты не существует.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                //
                if (scan.Length == 5)
                {
                    bool log = true;
                    for (int i = 0; i < dataGridView8.RowCount; i++)
                    {
                        string white_id = dataGridView8.Rows[i].Cells[1].Value.ToString();
                        if (white_id == scan)
                        {
                            dataGridView8.CurrentCell = dataGridView8.Rows[i].Cells[0];
                            dataGridView8.Rows[i].Selected = true;
                            log = false;
                            break;
                        }
                    }
                    if (log == true) {
                        MessageBox.Show("Указанная коробка не привязана к данному контракту.", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch { }
            comboBox3.Text = label12.Text;
            timer3.Enabled = false;
            //
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            timer3.Enabled = true;
            timer_focus.Enabled = true;
        }

        int _palet_add = 0;
        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView9.CurrentCell.RowIndex;
                string pid = dataGridView9.Rows[selected].Cells[1].Value.ToString();
                if (_palet_add == 0)
                {
                    GetPaletBoxes(pid, "", 0);
                }
                else
                {
                    GetPaletBoxes(pid, label12.Text, 1);
                    textBox11.Text = pid;
                }
                textBox10.Focus();
                comboBox3.Text = label12.Text;
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text != "") {
                _palet_add = 1;
                button13.Enabled = true;
                button14.Enabled = true;
                button12.Enabled = true;
                button15.Enabled = true;
                button16.Enabled = false;
                button17.Enabled = true;
                comboBox3.Enabled = false;
                textBox11.Text = GeneratePaletID("");
                int selected = 0;
                try
                {
                    selected = dataGridView9.CurrentCell.RowIndex;
                    string pid = dataGridView9.Rows[selected].Cells[1].Value.ToString();
                    GetPaletBoxes(pid, label12.Text, 1);
                }
                catch { }
                textBox10.Enabled = true;
                textBox10.Focus();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            _palet_add = 0;
            button13.Enabled = false;
            button14.Enabled = false;
            button12.Enabled = false;
            button15.Enabled = false;
            button16.Enabled = true;
            button17.Enabled = false;
            comboBox3.Enabled = true;
            textBox11.Clear();
            int selected = 0;
            try
            {
                selected = dataGridView9.CurrentCell.RowIndex;
                string pid = dataGridView9.Rows[selected].Cells[1].Value.ToString();
                GetPaletBoxes(pid, "", 0);
            }
            catch { }
            textBox10.Enabled = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            _palet_add = 0;
            button13.Enabled = false;
            button14.Enabled = false;
            button12.Enabled = false;
            button15.Enabled = false;
            button16.Enabled = true;
            button17.Enabled = false;
            comboBox3.Enabled = true;
            textBox11.Clear();
            int selected = 0;
            try
            {
                selected = dataGridView9.CurrentCell.RowIndex;
                string pid = dataGridView9.Rows[selected].Cells[1].Value.ToString();
                GetPaletBoxes(pid, "", 0);
            }
            catch { }
            textBox10.Enabled = false;
        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            textBox11.Text = GeneratePaletID("");
            textBox10.Focus();
        }

        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            textBox10.Focus();
        }

        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            comboBox3.Text = label12.Text;
            textBox10.Focus();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            int selectedRowIndex = 0;
            selectedRowIndex = dataGridView9.CurrentCell.RowIndex;
            string poleta = dataGridView9.Rows[selectedRowIndex].Cells[1].Value.ToString();
            GetListOfBoxes(poleta);
        }
        public class ListOfBoxes
        {
            public string contractId { get; set; }
            public string boxNo { get; set; }
            public string clientOrderId { get; set; }
        }
        public class CheckListContractId
        {
            public string contractId { get; set; }
        }
        public void GetListOfBoxes(string polet_id)
        {
            List<ListOfBoxes> boxes = new List<ListOfBoxes>();
            string poletNo = "";
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT DISTINCT collected_contract_items.white_id, collected_contract_items.id_contract, collected_contract_items.no_poleta, collected_contracts.client_order_id FROM collected_contracts join collected_contract_items ON collected_contracts.id_contract=collected_contract_items.id_contract WHERE no_poleta=@polet ORDER BY white_id";
            comm.Parameters.AddWithValue("@polet", polet_id);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                poletNo = reader.GetValue(2).ToString();
                ListOfBoxes box = new ListOfBoxes();
                box.contractId = reader.GetValue(1).ToString();
                box.boxNo = reader.GetValue(0).ToString();
                box.clientOrderId = reader.GetValue(3).ToString();
                boxes.Add(box);
            }
            reader.Close();

            GenerateListOfBoxes(boxes, poletNo);
        }
        public void GenerateListOfBoxes(List<ListOfBoxes> listOfBoxes, string polet_id)
        {
            List<CheckListContractId> checkListContract = new List<CheckListContractId>();
            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");
            worksheet.Column(1).Width = 4;
            worksheet.Column(2).Width = 8;
            worksheet.Column(3).Width = 14;
            worksheet.Column(4).Width = 7;
            worksheet.Column(5).Width = 8;
            worksheet.Column(6).Width = 15;
            worksheet.Column(7).Width = 9;
            worksheet.Column(8).Width = 9;

            int rowCount = 1;
            string contractId_check = "";
            for (int i = 0; i < listOfBoxes.Count; i++)
            {
                bool check = true;
                for(int m=0; m<checkListContract.Count; m++)
                {
                    if(checkListContract[m].contractId == listOfBoxes[i].contractId)
                    {
                        check = false;
                    }
                }

                if (check == true)
                {
                    CheckListContractId contractId = new CheckListContractId();
                    contractId.contractId = listOfBoxes[i].contractId;
                    checkListContract.Add(contractId);

                    worksheet.Cells[rowCount, 1].Value = "Contract:";
                    worksheet.Cells[rowCount, 1, rowCount, 2].Merge = true;
                    worksheet.Cells[rowCount, 3].Value = listOfBoxes[i].contractId;
                    worksheet.Cells[rowCount, 3, rowCount, 6].Merge = true;
                    worksheet.Cells[rowCount, 7].Value = "Order:";
                    worksheet.Cells[rowCount, 8].Value = listOfBoxes[i].clientOrderId;
                    worksheet.Cells[rowCount, 8, rowCount, 9].Merge = true;

                    using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
                    {
                        rng.Style.Font.Name = "Times New Roman";
                    }

                    rowCount++;
                    worksheet.Cells[rowCount, 1, rowCount, 2].Merge = true;
                    worksheet.Cells[rowCount, 3].Value = listOfBoxes[i].contractId;
                    worksheet.Cells[rowCount, 3, rowCount, 6].Merge = true;
                    worksheet.Cells[rowCount, 8].Value = listOfBoxes[i].clientOrderId;
                    worksheet.Cells[rowCount, 8, rowCount, 9].Merge = true;

                    using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
                    {
                        rng.Style.Font.Name = "IDAutomationHC39M";
                    }

                    contractId_check = listOfBoxes[i].contractId;
                    rowCount++;
                }
            }


            worksheet.Cells[rowCount, 1].Value = "Id Poleta:";
            worksheet.Cells[rowCount, 1, rowCount, 2].Merge = true;
            worksheet.Cells[rowCount, 3].Value = polet_id;
            worksheet.Cells[rowCount, 3, rowCount, 6].Merge = true;
            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 6])
            {
                rng.Style.Font.Name = "Times New Roman";
            }

            rowCount++;

            worksheet.Cells[rowCount, 1, rowCount, 2].Merge = true;
            worksheet.Cells[rowCount, 3].Value = polet_id;
            worksheet.Cells[rowCount, 3, rowCount, 6].Merge = true;
            worksheet.Cells[rowCount, 3].Style.Font.Name = "IDAutomationHC39M";

            rowCount++;

            worksheet.Cells[rowCount, 1].Value = "#";
            worksheet.Cells[rowCount, 2].Value = "Box Number";
            worksheet.Cells[rowCount, 2, rowCount, 6].Merge = true;

            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 6])
            {
                rng.Style.Font.Name = "Times New Roman";
            }
            rowCount++;
            int startBoxesRowIndex = rowCount;
            int boxCount = 0;
            int count = 0;
            for (count = 0; count < listOfBoxes.Count; count++)
            {
                boxCount++;
                worksheet.Cells[rowCount, 1].Value = boxCount;
                worksheet.Cells[rowCount, 2].Value = listOfBoxes[count].boxNo;
                worksheet.Cells[rowCount, 2, rowCount, 4].Merge = true;
                worksheet.Cells[rowCount, 5].Value = listOfBoxes[count].boxNo;
                worksheet.Cells[rowCount, 5, rowCount, 6].Merge = true;
                rowCount++;
            }


            worksheet.Cells[rowCount, 1].Value = "TOTAL: " + boxCount.ToString() + " boxes, ";
            worksheet.Cells[rowCount, 1, rowCount, 6].Merge = true;
            worksheet.Cells[rowCount, 1, rowCount, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowCount, 1, rowCount, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

            using (ExcelRange rng = worksheet.Cells[startBoxesRowIndex, 1, worksheet.Dimension.Rows, 4])
            {
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[startBoxesRowIndex, 5, worksheet.Dimension.Rows, 6])
            {
                rng.Style.Font.Name = "IDAutomationHC39M";
            }

            using (ExcelRange rng = worksheet.Cells[1, 1, startBoxesRowIndex - 4, 9])
            {
                rng.Style.WrapText = true;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Font.Size = 11;
            }
            using (ExcelRange rng = worksheet.Cells[startBoxesRowIndex - 3, 1, worksheet.Dimension.Rows, 6])
            {
                rng.Style.WrapText = true;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Font.Size = 11;
            }
            for (int k = 1; k <= worksheet.Dimension.Rows + 1; k++)
            {
                worksheet.Row(k).Height = 30;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                excel.SaveAs(new FileInfo(saveFileDialog.FileName));
            }
        }
    }
}
