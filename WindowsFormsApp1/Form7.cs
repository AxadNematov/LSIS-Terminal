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
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;

namespace WindowsFormsApp1
{
    public partial class Form7 : Form
    {
        private Form1 Frm1;
        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        string app_dir = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString());
        string app_dir_temp = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString()) + "\\Temp\\";

        public Form7(Form1 f1)
        {
            InitializeComponent();
            Frm1 = f1;
            InitGrids();
            FindCompany("");
        }

        private void Form7_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        public void InitGrids()
        {
            // Company List
            dataGridView4.RowCount = 0;
            dataGridView4.ColumnCount = 3;

            dataGridView4.Columns[0].HeaderText = "#";
            dataGridView4.Columns[1].HeaderText = "Название компании";
            dataGridView4.Columns[2].HeaderText = "db_id";

            dataGridView4.Columns[0].Width = 40;
            dataGridView4.Columns[1].Width = 260;
            dataGridView4.Columns[2].Width = 130;
            dataGridView4.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView4.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            //
            // Adres list
            dataGridView10.RowCount = 0;
            dataGridView10.ColumnCount = 3;

            dataGridView10.Columns[0].HeaderText = "#";
            dataGridView10.Columns[1].HeaderText = "Адрес";
            dataGridView10.Columns[2].HeaderText = "db_id";

            dataGridView10.Columns[0].Width = 40;
            dataGridView10.Columns[1].Width = 255;
            dataGridView10.Columns[2].Width = 130;
            dataGridView10.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView10.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            ///
            // Contacts list
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 3;

            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "Номер телефона";
            dataGridView1.Columns[2].HeaderText = "db_id";

            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 130;
            dataGridView1.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            ///
            // Port list
            dataGridView5.RowCount = 0;
            dataGridView5.ColumnCount = 4;

            dataGridView5.Columns[0].HeaderText = "#";
            dataGridView5.Columns[1].HeaderText = "Тип";
            dataGridView5.Columns[2].HeaderText = "Пункт назначения";
            dataGridView5.Columns[3].HeaderText = "db_id";

            dataGridView5.Columns[0].Width = 40;
            dataGridView5.Columns[1].Width = 50;
            dataGridView5.Columns[2].Width = 200;
            dataGridView5.Columns[3].Width = 150;
            dataGridView5.Columns[3].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView5.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            ///
            // Notify Company List
            dataGridView3.RowCount = 0;
            dataGridView3.ColumnCount = 3;

            dataGridView3.Columns[0].HeaderText = "#";
            dataGridView3.Columns[1].HeaderText = "Название компании";
            dataGridView3.Columns[2].HeaderText = "db_id";

            dataGridView3.Columns[0].Width = 40;
            dataGridView3.Columns[1].Width = 260;
            dataGridView3.Columns[2].Width = 130;
            dataGridView3.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView3.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            //
            // Notify Adres list
            dataGridView2.RowCount = 0;
            dataGridView2.ColumnCount = 3;

            dataGridView2.Columns[0].HeaderText = "#";
            dataGridView2.Columns[1].HeaderText = "Адрес";
            dataGridView2.Columns[2].HeaderText = "db_id";

            dataGridView2.Columns[0].Width = 40;
            dataGridView2.Columns[1].Width = 255;
            dataGridView2.Columns[2].Width = 130;
            dataGridView2.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView2.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            ///
            // Notify Contacts list
            dataGridView6.RowCount = 0;
            dataGridView6.ColumnCount = 3;

            dataGridView6.Columns[0].HeaderText = "#";
            dataGridView6.Columns[1].HeaderText = "Номер телефона";
            dataGridView6.Columns[2].HeaderText = "db_id";

            dataGridView6.Columns[0].Width = 40;
            dataGridView6.Columns[1].Width = 150;
            dataGridView6.Columns[2].Width = 130;
            dataGridView6.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView6.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            ///
        }

        public void AddCompany(string company_nm)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO crm_company (company) VALUES(@company_nm)";
            command1.Parameters.AddWithValue("@company_nm", company_nm);
            command1.ExecuteNonQuery();
            connection.Close();
        }

        public void UpdateCompany(string company_nm, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE crm_company SET company='" + company_nm + "' WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void DeleteCompany(string company_nm, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "DELETE FROM crm_company  WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
        }

        class CP_list
        {
            public int id { get; set; }
            public string company { get; set; }
        }

        public void FindCompany(string company_nm)
        {
            List<CP_list> c_list = new List<CP_list>();
            CP_list item = new CP_list();
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            //
            if (company_nm == "")
            {
                using (connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (command = new SqlCommand("SELECT * FROM crm_company ORDER BY id", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                item = new CP_list();
                                item.id = Convert.ToInt32(reader.GetValue(0).ToString());
                                item.company = reader.GetValue(1).ToString();
                                c_list.Add(item);
                            }
                        }
                    }
                    connection.Close();
                }
            }else{
                using (connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (command = new SqlCommand("SELECT * FROM crm_company WHERE company='" + company_nm + "' ORDER BY id", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                item = new CP_list();
                                item.id = Convert.ToInt32(reader.GetValue(0).ToString());
                                item.company = reader.GetValue(1).ToString();
                                c_list.Add(item);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            // Display
            dataGridView4.RowCount = c_list.Count;
            dataGridView4.RowHeadersWidth = 35;
            for (int i = 0; i < c_list.Count; i++)
            {
                dataGridView4.Rows[i].Cells[0].Value = i + 1;
                dataGridView4.Rows[i].Cells[1].Value = c_list[i].company;
                dataGridView4.Rows[i].Cells[2].Value = c_list[i].id;
                if (i % 2 == 0)
                {
                    dataGridView4.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView4.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView4.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView4.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView4.Rows[i].Cells[1].Style.BackColor = Color.White;
                    dataGridView4.Rows[i].Cells[2].Style.BackColor = Color.White;
                }
            }

        }

        private void button31_Click(object sender, EventArgs e)
        {
            if (textBox32.Text != "") {
                AddCompany(textBox32.Text);
                textBox32.Clear();
                FindCompany("");
                Frm1.FindCompanies();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            FindCompany(textBox32.Text);
        }

        private void label56_Click(object sender, EventArgs e)
        {
            textBox32.Clear();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            int selected = dataGridView4.CurrentCell.RowIndex;
            dataGridView4.ReadOnly = false;
            dataGridView4.CurrentCell = dataGridView4[1, selected];
            dataGridView4.BeginEdit(true);
        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int column = e.ColumnIndex;
            //
            string company_nm = dataGridView4.Rows[row].Cells[column].Value.ToString();
            int db_id = Convert.ToInt32(dataGridView4.Rows[row].Cells[column + 1].Value);
            UpdateCompany(company_nm, db_id);
            //
            dataGridView4.ReadOnly = true;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            try {
                int selected = dataGridView4.CurrentCell.RowIndex;
                int id_db = Convert.ToInt32(dataGridView4.Rows[selected].Cells[2].Value.ToString());
                DeleteCompany("", id_db);
                FindCompany(textBox32.Text);
                Frm1.FindCompanies();
            }
            catch { }
            
        }

        ////////// Adres
        public void AddCompanyAdres(string adres, int c_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO crm_adres (adres, company_id) VALUES(@adres, @c_id)";
            command1.Parameters.AddWithValue("@adres", adres);
            command1.Parameters.AddWithValue("@c_id", c_id);
            command1.ExecuteNonQuery();
            connection.Close();
        }

        public void UpdateCompanyAdres(string adres, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE crm_adres SET adres='" + adres + "' WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void DeleteCompanyAdres(string company_nm, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "DELETE FROM crm_adres  WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
        }

        class CP_Adres_list
        {
            public int id { get; set; }
            public string adres { get; set; }
        }

        public void FindCompanyAdres(int company_id)
        {
            List<CP_Adres_list> adres_list = new List<CP_Adres_list>();
            CP_Adres_list item = new CP_Adres_list();
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            //
            using (connection = new SqlConnection(conString))
            {
                connection.Open();
                using (command = new SqlCommand("SELECT * FROM crm_adres WHERE company_id='" + company_id + "' ORDER BY id", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            item = new CP_Adres_list();
                            item.id = Convert.ToInt32(reader.GetValue(0).ToString());
                            item.adres = reader.GetValue(1).ToString();
                            adres_list.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            // Display
            dataGridView10.RowCount = adres_list.Count;
            dataGridView10.RowHeadersWidth = 35;
            for (int i = 0; i < adres_list.Count; i++)
            {
                dataGridView10.Rows[i].Cells[0].Value = i + 1;
                dataGridView10.Rows[i].Cells[1].Value = adres_list[i].adres;
                dataGridView10.Rows[i].Cells[2].Value = adres_list[i].id;
                if (i % 2 == 0)
                {
                    dataGridView10.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView10.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView10.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView10.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView10.Rows[i].Cells[1].Style.BackColor = Color.White;
                    dataGridView10.Rows[i].Cells[2].Style.BackColor = Color.White;
                }
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                int selected = dataGridView4.CurrentCell.RowIndex;
                int c_id = Convert.ToInt32(dataGridView4.Rows[selected].Cells[2].Value.ToString());
                //
                AddCompanyAdres(textBox1.Text, c_id);
                textBox1.Clear();
                //
                FindCompanyAdres(c_id);
            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            try {
                dataGridView10.RowCount = 0;
                dataGridView1.RowCount = 0;
                dataGridView5.RowCount = 0;
                dataGridView3.RowCount = 0;
                dataGridView2.RowCount = 0;
                dataGridView6.RowCount = 0;
                //
                int selected = dataGridView4.CurrentCell.RowIndex;
                int c_id = Convert.ToInt32(dataGridView4.Rows[selected].Cells[2].Value.ToString());
                //
                FindCompanyAdres(c_id);
                FindCompanyDest(c_id);
            }
            catch{ }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int selected = dataGridView10.CurrentCell.RowIndex;
            dataGridView10.ReadOnly = false;
            dataGridView10.CurrentCell = dataGridView10[1, selected];
            dataGridView10.BeginEdit(true);
        }

        private void dataGridView10_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int column = e.ColumnIndex;
            //
            string adres = dataGridView10.Rows[row].Cells[column].Value.ToString();
            int db_id = Convert.ToInt32(dataGridView4.Rows[row].Cells[column + 1].Value);
            UpdateCompanyAdres(adres, db_id);
            //
            dataGridView10.ReadOnly = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                int selected = dataGridView10.CurrentCell.RowIndex;
                int id_db = Convert.ToInt32(dataGridView10.Rows[selected].Cells[2].Value.ToString());
                //
                DeleteCompanyAdres("", id_db);
                //
                selected = dataGridView4.CurrentCell.RowIndex;
                int c_id = Convert.ToInt32(dataGridView4.Rows[selected].Cells[2].Value.ToString());
                //
                FindCompanyAdres(c_id);
            }
            catch { }
        }

        private void label6_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        ////////// Phones
        public void AddCompanyPhones(string phone, int adres_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO crm_phones (phone, adres_id) VALUES(@phone, @adr_id)";
            command1.Parameters.AddWithValue("@phone", phone);
            command1.Parameters.AddWithValue("@adr_id", adres_id);
            command1.ExecuteNonQuery();
            connection.Close();
        }

        public void UpdateCompanyPhones(string phone, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE crm_phones SET phone='" + phone + "' WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void DeleteCompanyPhones(string company_nm, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "DELETE FROM crm_phones  WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
        }

        class CP_Phone_list
        {
            public int id { get; set; }
            public string number { get; set; }
        }

        public void FindCompanyPhones(int adres_id)
        {
            List<CP_Phone_list> phone_list = new List<CP_Phone_list>();
            CP_Phone_list item = new CP_Phone_list();
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            //
            using (connection = new SqlConnection(conString))
            {
                connection.Open();
                using (command = new SqlCommand("SELECT * FROM crm_phones WHERE adres_id='" + adres_id + "' ORDER BY id", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            item = new CP_Phone_list();
                            item.id = Convert.ToInt32(reader.GetValue(0).ToString());
                            item.number = reader.GetValue(1).ToString();
                            phone_list.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            // Display
            dataGridView1.RowCount = phone_list.Count;
            dataGridView1.RowHeadersWidth = 35;
            for (int i = 0; i < phone_list.Count; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = i + 1;
                dataGridView1.Rows[i].Cells[1].Value = phone_list[i].number;
                dataGridView1.Rows[i].Cells[2].Value = phone_list[i].id;
                if (i % 2 == 0)
                {
                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.White;
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.White;
                }
            }

        }

        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int selected = dataGridView10.CurrentCell.RowIndex;
                int adr_id = Convert.ToInt32(dataGridView10.Rows[selected].Cells[2].Value.ToString());
                //
                FindCompanyPhones(adr_id);
            }
            catch { }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                int selected = dataGridView10.CurrentCell.RowIndex;
                int adr_id = Convert.ToInt32(dataGridView10.Rows[selected].Cells[2].Value.ToString());
                //
                AddCompanyPhones(textBox2.Text, adr_id);
                textBox2.Clear();
                //
                FindCompanyPhones(adr_id);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int selected = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.ReadOnly = false;
            dataGridView1.CurrentCell = dataGridView1[1, selected];
            dataGridView1.BeginEdit(true);
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int column = e.ColumnIndex;
            //
            string phone = dataGridView1.Rows[row].Cells[column].Value.ToString();
            int db_id = Convert.ToInt32(dataGridView1.Rows[row].Cells[column + 1].Value);
            UpdateCompanyPhones(phone, db_id);
            //
            dataGridView1.ReadOnly = true;
        }

        private void label7_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
        }

        /////////////// Destination point
        private void checkBox1_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
            checkBox2.Checked = false;
        }

        private void checkBox2_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = true;
        }

        public void AddCompanyDestPoint(string type, string dest, int c_id)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO crm_destination (type, dest_point, company_id) VALUES(@type, @dest, @c_id)";
            command1.Parameters.AddWithValue("@type", type);
            command1.Parameters.AddWithValue("@dest", dest);
            command1.Parameters.AddWithValue("@c_id", c_id);
            command1.ExecuteNonQuery();
            connection.Close();
        }

        public void UpdateCompanyDestPoint(string dest, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE crm_destination SET dest_point='" + dest + "' WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void DeleteCompanyDestPoint(string dest, int id_db)
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "DELETE FROM crm_destination  WHERE id='" + id_db + "'";
            command.ExecuteNonQuery();
        }

        class CP_Dest_list
        {
            public int id { get; set; }
            public string type { get; set; }
            public string dest { get; set; }
        }

        public void FindCompanyDest(int company_id)
        {
            List<CP_Dest_list> dest_list = new List<CP_Dest_list>();
            CP_Dest_list item = new CP_Dest_list();
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            //
            using (connection = new SqlConnection(conString))
            {
                connection.Open();
                using (command = new SqlCommand("SELECT * FROM crm_destination WHERE company_id='" + company_id + "' ORDER BY id", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            item = new CP_Dest_list();
                            item.id = Convert.ToInt32(reader.GetValue(0).ToString());
                            item.type = reader.GetValue(1).ToString();
                            item.dest = reader.GetValue(2).ToString();
                            dest_list.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            // Display
            dataGridView5.RowCount = dest_list.Count;
            dataGridView5.RowHeadersWidth = 35;
            for (int i = 0; i < dest_list.Count; i++)
            {
                dataGridView5.Rows[i].Cells[0].Value = i + 1;
                dataGridView5.Rows[i].Cells[1].Value = dest_list[i].type;
                dataGridView5.Rows[i].Cells[2].Value = dest_list[i].dest;
                dataGridView5.Rows[i].Cells[3].Value = dest_list[i].id;
                if (i % 2 == 0)
                {
                    dataGridView5.Rows[i].Cells[0].Style.BackColor = Color.WhiteSmoke;
                    dataGridView5.Rows[i].Cells[1].Style.BackColor = Color.WhiteSmoke;
                    dataGridView5.Rows[i].Cells[2].Style.BackColor = Color.WhiteSmoke;
                    dataGridView5.Rows[i].Cells[3].Style.BackColor = Color.WhiteSmoke;
                }
                else
                {
                    dataGridView5.Rows[i].Cells[0].Style.BackColor = Color.White;
                    dataGridView5.Rows[i].Cells[1].Style.BackColor = Color.White;
                    dataGridView5.Rows[i].Cells[2].Style.BackColor = Color.White;
                    dataGridView5.Rows[i].Cells[3].Style.BackColor = Color.White;
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (textBox7.Text != "")
            {
                int selected = dataGridView4.CurrentCell.RowIndex;
                int c_id = Convert.ToInt32(dataGridView4.Rows[selected].Cells[2].Value.ToString());
                //
                string type = "";
                if (checkBox1.Checked == true) { type = "RW"; }
                if (checkBox2.Checked == true) { type = "AIR"; }
                AddCompanyDestPoint(type, textBox7.Text, c_id);
                textBox7.Clear();
                //
                FindCompanyDest(c_id);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int selected = dataGridView5.CurrentCell.RowIndex;
            dataGridView5.ReadOnly = false;
            dataGridView5.CurrentCell = dataGridView5[2, selected];
            dataGridView5.BeginEdit(true);
        }

        private void dataGridView5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int column = e.ColumnIndex;
            //
            string dest = dataGridView5.Rows[row].Cells[column].Value.ToString();
            int db_id = Convert.ToInt32(dataGridView5.Rows[row].Cells[column + 1].Value);
            UpdateCompanyDestPoint(dest, db_id);
            //
            dataGridView5.ReadOnly = true;
        }

        private void label8_Click(object sender, EventArgs e)
        {
            textBox7.Clear();
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                int selected = dataGridView5.CurrentCell.RowIndex;
                int id_db = Convert.ToInt32(dataGridView5.Rows[selected].Cells[3].Value.ToString());
                //
                DeleteCompanyDestPoint("", id_db);
            }
            catch { }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                int selected = dataGridView1.CurrentCell.RowIndex;
                int id_db = Convert.ToInt32(dataGridView1.Rows[selected].Cells[2].Value.ToString());
                //
                DeleteCompanyPhones("", id_db);
                //
                selected = dataGridView10.CurrentCell.RowIndex;
                int adr_id = Convert.ToInt32(dataGridView10.Rows[selected].Cells[2].Value.ToString());
                //
                FindCompanyPhones(adr_id);
            }
            catch { }
        }
    }
}
