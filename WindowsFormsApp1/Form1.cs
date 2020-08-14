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
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Runtime.InteropServices;
using System.Threading;
using OfficeOpenXml;
using System.Globalization;
using System.Drawing.Drawing2D;
using System.Speech.Recognition;
using System.Speech.Synthesis;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        static SpeechRecognitionEngine engine;
        public int _current_shelf_X = -1;
        public int _current_shelf_Y = -1;
        public int _current_shelf_Z = -1;

        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        string app_dir = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString());
        string app_dir_temp = System.IO.Path.GetDirectoryName(Application.ExecutablePath.ToString()) + "\\Temp\\";

        public string shelf_mode = "search";

        public Form2 Frm2;
        public Form3 Frm3;
        public Form4 Frm4;
        public Form5 Frm5;
        public Form6 Frm6;
        public Form7 Frm7;

        Color cl_odd = Color.FromArgb(255, 243, 240, 240);
        Color cl_even = Color.FromArgb(255, 255, 255, 255);

        Color cl_hc1 = Color.FromArgb(255, 82, 102, 107);
        Color cl_hc2 = Color.FromArgb(255, 29, 51, 60);
        Color cl_hc3 = Color.FromArgb(255, 82, 102, 107);

        Color cl_cp_header = Color.FromArgb(255, 237, 237, 237);

        public Form1()
        {
            InitializeComponent();
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            Frm2 = new Form2(this);
            Frm2.TopMost = true;
            Frm3 = new Form3(this);
            Frm3.TopMost = true;
            Frm4 = new Form4();
            Frm4.TopMost = true;
            Frm5 = new Form5(this);
            Frm5.TopMost = true;
            Frm6 = new Form6(this);
            Frm6.TopMost = true;
            Frm7 = new Form7(this);
            Frm7.TopMost = true;

            //engine = new SpeechRecognitionEngine();
            //engine.SetInputToDefaultAudioDevice();
            //engine.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(new String[] { "Hello", "привет", "открыть меню", "закрыть меню", "How are you", "interesting", "Open menu", "Close menu" }))));
            //engine.RecognizeAsync(RecognizeMode.Multiple);
            //engine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(engine_SpeechRecognixed);
        }
        private void engine_SpeechRecognixed(object sender, SpeechRecognizedEventArgs e)
        {
            string result = e.Result.Text;
            if (result == "Hello")
            {
                SpeechSynthesizer audio = new SpeechSynthesizer();
                audio.Speak("Hi");
            }
            if (result == "привет")
            {
                SpeechSynthesizer audio = new SpeechSynthesizer();
                audio.Speak("Здравствуйте");
            }
            if (result == "Open menu" || result == "открыть меню")
            {
                if (panel_main_menu.Width < 350)
                {
                    Size sz = new Size();
                    sz.Width = 350;
                    panel_main_menu.Width = sz.Width;
                }
                else
                {
                    SpeechSynthesizer audio = new SpeechSynthesizer();
                    audio.Speak("It is already open");
                }
            }
            if (result == "Close menu" || result == "закрыть меню")
            {
                if (panel_main_menu.Width == 350)
                {
                    Size sz = new Size();
                    sz.Width = 65;
                    panel_main_menu.Width = sz.Width;
                }
                else
                {
                    SpeechSynthesizer audio = new SpeechSynthesizer();
                    audio.Speak("It is closed");
                }
            }
        }
        public void ShowPanels(string show)
        {
            //panel_Main_frame.Dock = DockStyle.None;
            //Point p = new Point();
            //p.Y = -6000;
            //panel_Main_frame.Location = p; 
            // List of panels to hide
            //...
            // Panel to show

            Size ns = new Size();
            ns.Width = this.Width;
            ns.Height = this.Height;
            //
            Point np = new Point();
            np.Y = 0;
            if (panel_main_menu.Width == 350) {
                np.X = 352;
            }
            else {
                np.X = 67;
            }
            //
            panel_cover.Size = ns;
            panel_cover.Location = np;
            //
            if (show == "Заказ") {
                panel_shelfs_control.Dock = DockStyle.None;
                panel_shelfs_control.Visible = false;
                panel_income.Dock = DockStyle.None;
                panel_income.Visible = false;
                panel_com_proposal.Dock = DockStyle.None;
                panel_com_proposal.Visible = false;
                panel_basket.Dock = DockStyle.None;
                panel_basket.Visible = false;
                //
                panel_make_order.Dock = DockStyle.Fill;
                panel_make_order.Visible = true;
            }
            if (show == "Стелажи")
            {
                panel_make_order.Dock = DockStyle.None;
                panel_make_order.Visible = false;
                panel_income.Dock = DockStyle.None;
                panel_income.Visible = false;
                panel_com_proposal.Dock = DockStyle.None;
                panel_com_proposal.Visible = false;
                panel_basket.Dock = DockStyle.None;
                panel_basket.Visible = false;
                //
                panel_shelfs_control.Dock = DockStyle.Fill;
                panel_shelfs_control.Visible = true;
            }
            if (show == "Приход")
            {
                panel_make_order.Dock = DockStyle.None;
                panel_make_order.Visible = false;
                panel_shelfs_control.Dock = DockStyle.None;
                panel_shelfs_control.Visible = false;
                panel_com_proposal.Dock = DockStyle.None;
                panel_com_proposal.Visible = false;
                panel_basket.Dock = DockStyle.None;
                panel_basket.Visible = false;
                //
                panel_income.Dock = DockStyle.Fill;
                panel_income.Visible = true;
            }
            if (show == "Ком")
            {
                panel_make_order.Dock = DockStyle.None;
                panel_make_order.Visible = false;
                panel_shelfs_control.Dock = DockStyle.None;
                panel_shelfs_control.Visible = false;
                panel_income.Dock = DockStyle.None;
                panel_income.Visible = false;
                panel_basket.Dock = DockStyle.None;
                panel_basket.Visible = false;
                //
                panel_com_proposal.Dock = DockStyle.Fill;
                panel_com_proposal.Visible = true;
            }
            if (show == "Корзина")
            {
                panel_make_order.Dock = DockStyle.None;
                panel_make_order.Visible = false;
                panel_shelfs_control.Dock = DockStyle.None;
                panel_shelfs_control.Visible = false;
                panel_income.Dock = DockStyle.None;
                panel_income.Visible = false;
                panel_com_proposal.Dock = DockStyle.None;
                panel_com_proposal.Visible = false;
                //
                panel_basket.Dock = DockStyle.Fill;
                panel_basket.Visible = true;
            }
            //
            //panel_Main_frame.Dock = DockStyle.Fill;
            timer_panel_cover.Enabled = true;
        }

        public void ClearCash()
        {
            var files = Directory.GetFiles(app_dir_temp);
            for (int i = 0; i < files.Count(); i++)
            {
                string s = files[i].Substring(files[i].Length - 16, 15);
                if (s.IndexOf("temp") != -1)
                {
                    try
                    {
                        File.Delete(files[i]);
                    }
                    catch { }
                }
            }

        }

        // CP history
        public void InitGrid10()
        {
            // CPs
            dataGridView10.RowCount = 0;
            dataGridView10.ColumnCount = 4;

            dataGridView10.Columns[0].HeaderText = "#";
            dataGridView10.Columns[1].HeaderText = "CP ID";
            dataGridView10.Columns[2].HeaderText = "Заказчик";
            dataGridView10.Columns[3].HeaderText = "Проект";

            dataGridView10.Columns[0].Width = 40;
            dataGridView10.Columns[1].Width = 100;
            dataGridView10.Columns[2].Width = 130;
            dataGridView10.Columns[3].Width = 130;
            //
            foreach (DataGridViewColumn col in dataGridView10.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.ForeColor = Color.White;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            // CP versions
            dataGridView11.RowCount = 0;
            dataGridView11.ColumnCount = 9;

            dataGridView11.Columns[0].HeaderText = "Версия";
            dataGridView11.Columns[1].HeaderText = "Дата";
            dataGridView11.Columns[2].HeaderText = "Доставка";
            dataGridView11.Columns[3].HeaderText = "Order ID";
            dataGridView11.Columns[4].HeaderText = "Order Дата";
            dataGridView11.Columns[5].HeaderText = "ID контракта";
            dataGridView11.Columns[6].HeaderText = "Дата контракта";
            dataGridView11.Columns[7].HeaderText = "PO ID";
            dataGridView11.Columns[8].HeaderText = "PO Дата";

            dataGridView11.Columns[0].Width = 50;
            dataGridView11.Columns[1].Width = 75;
            dataGridView11.Columns[2].Width = 70;
            dataGridView11.Columns[3].Width = 85;
            dataGridView11.Columns[4].Width = 80;
            dataGridView11.Columns[5].Width = 110;
            dataGridView11.Columns[6].Width = 100;
            dataGridView11.Columns[7].Width = 85;
            dataGridView11.Columns[8].Width = 70;

            //
            foreach (DataGridViewColumn col in dataGridView11.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }

            // axad 22.07 start
            dataGridView13.RowCount = 0;
            dataGridView13.ColumnCount = 4;

            dataGridView13.Columns[0].HeaderText = "#";
            dataGridView13.Columns[1].HeaderText = "Id Контракта";
            dataGridView13.Columns[2].HeaderText = "Заказчик";
            dataGridView13.Columns[3].HeaderText = "Сумма";


            dataGridView13.Columns[0].Width = 40;
            dataGridView13.Columns[1].Width = 120;
            dataGridView13.Columns[2].Width = 100;
            dataGridView13.Columns[3].Width = 80;

            dataGridView24.RowCount = 0;
            dataGridView24.ColumnCount = 4;

            dataGridView24.Columns[0].HeaderText = "Версия";
            dataGridView24.Columns[1].HeaderText = "Доставка";
            dataGridView24.Columns[2].HeaderText = "Пункт доставки";
            dataGridView24.Columns[3].HeaderText = "Дата";

            dataGridView24.Columns[0].Width = 60;
            dataGridView24.Columns[1].Width = 60;
            dataGridView24.Columns[2].Width = 380;
            dataGridView24.Columns[3].Width = 100;

            // axad 22.07 end

            foreach (DataGridViewColumn col in dataGridView13.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            dataGridView14.RowCount = 0;
            dataGridView14.ColumnCount = 6;

            dataGridView14.Columns[0].HeaderText = "#";
            dataGridView14.Columns[1].HeaderText = "Код продукта";
            dataGridView14.Columns[2].HeaderText = "Название продукта";
            dataGridView14.Columns[3].HeaderText = "Кол-во";
            dataGridView14.Columns[4].HeaderText = "Цена за ед";
            dataGridView14.Columns[5].HeaderText = "Общая сумма";

            dataGridView14.Columns[0].Width = 40;
            dataGridView14.Columns[1].Width = 100;
            dataGridView14.Columns[2].Width = 250;
            dataGridView14.Columns[3].Width = 70;
            dataGridView14.Columns[4].Width = 90;
            dataGridView14.Columns[5].Width = 100;


            foreach (DataGridViewColumn col in dataGridView14.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            // Templates
            dataGridView22.RowCount = 0;
            dataGridView22.ColumnCount = 3;

            dataGridView22.Columns[0].HeaderText = "#";
            dataGridView22.Columns[1].HeaderText = "Шаблон";
            dataGridView22.Columns[2].HeaderText = "";


            dataGridView22.Columns[0].Width = 40;
            dataGridView22.Columns[1].Width = 100;
            dataGridView22.Columns[2].Width = 100;

            dataGridView22.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView22.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
            // Templates History
            dataGridView23.RowCount = 0;
            dataGridView23.ColumnCount = 3;

            dataGridView23.Columns[0].HeaderText = "#";
            dataGridView23.Columns[1].HeaderText = "Шаблон";
            dataGridView23.Columns[2].HeaderText = "";


            dataGridView23.Columns[0].Width = 40;
            dataGridView23.Columns[1].Width = 100;
            dataGridView23.Columns[2].Width = 100;

            dataGridView23.Columns[2].Visible = false;

            //
            foreach (DataGridViewColumn col in dataGridView23.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
        }

        public class CPs
        {
            public string cp_id { get; set; }
            public string cp_date { get; set; }
            public string delivery { get; set; }
            public string version { get; set; }
            public string po_id { get; set; }
            public string po_date { get; set; }
            public string order_id { get; set; }
            public string order_date { get; set; }
            public string ct_id { get; set; }
            public string ct_date { get; set; }
            public string customer { get; set; }
            public string project { get; set; }
        }

        class CP_list
        {
            public int id { get; set; }
            public string company { get; set; }
        }

        public void FindCompanies()
        {
            List<CP_list> c_list = new List<CP_list>();
            CP_list item = new CP_list();
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            //
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
            comboBox14.Items.Clear();
            comboBox15.Items.Clear();
            // Display
            for (int i = 0; i < c_list.Count; i++) {
                comboBox14.Items.Add(c_list[i].company);
                comboBox15.Items.Add(c_list[i].id);
            }
            comboBox14.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
        }

        public void GetCPList(string argument, int type)
        {
            // search argument cp_id
            if (type == 0)
            {
                if (argument == "")
                {
                    List<CPs> cp_items = new List<CPs>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT cp_id, customer, project, id FROM com_proposal WHERE version=1 ORDER BY id DESC", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    CPs item = new CPs();
                                    item.cp_id = reader.GetValue(0).ToString();
                                    item.customer = reader.GetValue(1).ToString();
                                    item.project = reader.GetValue(2).ToString();
                                    cp_items.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView10.RowCount = 0;
                    dataGridView11.RowCount = 0;
                    dataGridView10.RowCount = cp_items.Count;
                    dataGridView10.RowHeadersWidth = 35;
                    for (int i = 0; i < cp_items.Count; i++)
                    {
                        dataGridView10.Rows[i].Cells[0].Value = i + 1;
                        dataGridView10.Rows[i].Cells[1].Value = cp_items[i].cp_id;
                        dataGridView10.Rows[i].Cells[2].Value = cp_items[i].customer;
                        dataGridView10.Rows[i].Cells[3].Value = cp_items[i].project;
                        if (i % 2 == 0)
                        {
                            dataGridView10.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[2].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView10.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[2].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
                //
                if (argument != "")
                {
                    List<CPs> cp_items = new List<CPs>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT cp_id, customer, project, id FROM com_proposal WHERE version=1 AND cp_id='" + argument + "' ORDER BY id DESC", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    CPs item = new CPs();
                                    item.cp_id = reader.GetValue(0).ToString();
                                    item.customer = reader.GetValue(1).ToString();
                                    item.project = reader.GetValue(2).ToString();
                                    cp_items.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView10.RowCount = 0;
                    dataGridView11.RowCount = 0;
                    dataGridView10.RowCount = cp_items.Count;
                    dataGridView10.RowHeadersWidth = 35;
                    for (int i = 0; i < cp_items.Count; i++)
                    {
                        dataGridView10.Rows[i].Cells[0].Value = i + 1;
                        dataGridView10.Rows[i].Cells[1].Value = cp_items[i].cp_id;
                        dataGridView10.Rows[i].Cells[2].Value = cp_items[i].customer;
                        dataGridView10.Rows[i].Cells[3].Value = cp_items[i].project;
                        if (i % 2 == 0)
                        {
                            dataGridView10.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[2].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView10.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[2].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
            // search argument date
            if (type == 1)
            {
                //
                if (argument != "")
                {
                    List<CPs> cp_items = new List<CPs>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT cp_id, customer, project, id FROM com_proposal WHERE version=1 AND cp_date='" + argument + "' ORDER BY id DESC", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    CPs item = new CPs();
                                    item.cp_id = reader.GetValue(0).ToString();
                                    item.customer = reader.GetValue(1).ToString();
                                    item.project = reader.GetValue(2).ToString();
                                    cp_items.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView10.RowCount = 0;
                    dataGridView11.RowCount = 0;
                    dataGridView10.RowCount = cp_items.Count;
                    dataGridView10.RowHeadersWidth = 35;
                    for (int i = 0; i < cp_items.Count; i++)
                    {
                        dataGridView10.Rows[i].Cells[0].Value = i + 1;
                        dataGridView10.Rows[i].Cells[1].Value = cp_items[i].cp_id;
                        dataGridView10.Rows[i].Cells[2].Value = cp_items[i].customer;
                        dataGridView10.Rows[i].Cells[3].Value = cp_items[i].project;
                        if (i % 2 == 0)
                        {
                            dataGridView10.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[2].Style.BackColor = cl_even;
                            dataGridView10.Rows[i].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView10.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[2].Style.BackColor = cl_odd;
                            dataGridView10.Rows[i].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
        }

        public void GetCPVersions(string cp_id)
        {
            List<CPs> cp_items = new List<CPs>();

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT rw, air FROM com_proposal WHERE cp_id='" + cp_id + "' and version=" + 1, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            textBox22.Text = reader.GetValue(0).ToString();
                            textBox21.Text = reader.GetValue(1).ToString();
                        }
                    }
                }
                connection.Close();
            }

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT version, cp_date, delivery_type, po_id, po_date, order_id, order_date, ct_id, ct_date FROM com_proposal WHERE cp_id='" + cp_id + "' ORDER BY version", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CPs item = new CPs();
                            item.version = reader.GetValue(0).ToString();
                            item.cp_date = reader.GetValue(1).ToString();
                            item.delivery = reader.GetValue(2).ToString();
                            //
                            item.po_id = reader.GetValue(3).ToString();
                            item.po_date = reader.GetValue(4).ToString();
                            item.order_id = reader.GetValue(5).ToString();
                            item.order_date = reader.GetValue(6).ToString();
                            item.ct_id = reader.GetValue(7).ToString();
                            item.ct_date = reader.GetValue(8).ToString();
                            cp_items.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            dataGridView11.RowCount = 0;
            dataGridView11.RowCount = cp_items.Count;
            dataGridView11.RowHeadersWidth = 35;
            for (int i = 0; i < cp_items.Count; i++)
            {
                dataGridView11.Rows[i].Cells[0].Value = cp_items[i].version;
                dataGridView11.Rows[i].Cells[1].Value = cp_items[i].cp_date;
                dataGridView11.Rows[i].Cells[2].Value = cp_items[i].delivery;
                //
                dataGridView11.Rows[i].Cells[3].Value = cp_items[i].order_id;
                dataGridView11.Rows[i].Cells[4].Value = cp_items[i].order_date;
                dataGridView11.Rows[i].Cells[5].Value = cp_items[i].ct_id;
                dataGridView11.Rows[i].Cells[6].Value = cp_items[i].ct_date;
                dataGridView11.Rows[i].Cells[7].Value = cp_items[i].po_id;
                dataGridView11.Rows[i].Cells[8].Value = cp_items[i].po_date;

                if (i % 2 == 0)
                {
                    dataGridView11.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[5].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[6].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[7].Style.BackColor = cl_even;
                    dataGridView11.Rows[i].Cells[8].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView11.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[5].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[6].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[7].Style.BackColor = cl_odd;
                    dataGridView11.Rows[i].Cells[8].Style.BackColor = cl_odd;
                }
            }
        }

        public void GenerateCP(DataGridView Grid, ComboBox cBox, int type)
        {
            string s;
            s = comboBox16.Text;
            s = s.Trim();
            comboBox16.Text = s;
            //
            s = comboBox17.Text;
            s = s.Trim();
            comboBox17.Text = s;
            //
            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");
            worksheet.DefaultColWidth = 10;
            worksheet.Column(1).Width = 4;
            worksheet.Column(2).Width = 9;
            worksheet.Column(3).Width = 27;
            worksheet.Column(4).Width = 8;
            worksheet.Column(5).Width = 8;
            worksheet.Column(6).Width = 8;
            worksheet.Column(7).Width = 8;
            worksheet.Column(8).Width = 10;
            worksheet.Column(9).Width = 10;

            //using (System.Drawing.Image header = System.Drawing.Image.FromFile(app_dir_temp + "cp_header.png"))
            //{
            //    var excelImage = worksheet.Drawings.AddPicture("Header", header);
            //    excelImage.SetPosition(0, 0, 0, 0);
            //    if (cBox.SelectedIndex == 0)
            //    {
            //        excelImage.SetSize(825, 40);
            //    }
            //    else
            //    {
            //        excelImage.SetSize(620, 40);
            //    }

            //}

            using (System.Drawing.Image logo = System.Drawing.Image.FromFile(app_dir_temp + "Logo.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("My Logo", logo);
                if (cBox.SelectedIndex == 0)
                {
                    excelImage.SetPosition(0, 0, 2, 280);
                }
                else
                {
                    excelImage.SetPosition(0, 0, 2, 130);
                }
                excelImage.SetSize(25);
            }

            int rowCount = 14;
            int totalQty = 0;
            double totalAmountPriceRW = 0;
            double totalAmountPriceAIR = 0;
            int i = 0;

            worksheet.Cells[13, 1].Value = "#";
            worksheet.Cells[13, 2].Value = "Item code";
            worksheet.Cells[13, 3].Value = "Description of goods";
            worksheet.Cells[13, 4].Value = "HS code";
            worksheet.Cells[13, 5].Value = "Qty";
            if (cBox.SelectedIndex == 0)
            {
                worksheet.Row(13).Height = 75;
                worksheet.Cells[13, 6].Value = "Unit price in USD by RW, " + comboBox16.Text;
                worksheet.Cells[13, 7].Value = "Unit price in USD by AIR, " + comboBox17.Text;
                worksheet.Cells[13, 8].Value = "Amount price in USD by RW, " + comboBox16.Text;
                worksheet.Cells[13, 9].Value = "Amount price in USD by AIR, " + comboBox17.Text;
                worksheet.Cells[13, 10].Value = "PRODUCTION TIME";
                worksheet.Cells[13, 11].Value = "Delivery by AIR";
                worksheet.Cells[13, 12].Value = "Delivery by RW";

                if (textBox15.Text.Length > 0)
                {
                    worksheet.Cells[9, 10].Value = "CP:";
                    worksheet.Cells[9, 11].Value = textBox15.Text + "-1";
                    worksheet.Cells[9, 12].Value = dateTimePicker10.Value.Date.ToString("dd.MM.yyyy");

                }
                if (textBox30.Text.Length > 0)
                {
                    worksheet.Cells[11, 1].Value = "ORDER:     " + textBox30.Text + "     DATE:   " + dateTimePicker15.Value.ToShortDateString();
                    worksheet.Cells[11, 1, 11, 3].Merge = true;
                }

                if (textBox35.Text.Length > 0)
                {
                    worksheet.Cells[10, 1].Value = "PROJECT:     " + textBox35.Text;
                    worksheet.Cells[10, 1, 10, 3].Merge = true;
                }
                if (comboBox14.Text.Length > 0)
                {
                    worksheet.Cells[9, 1].Value = "TO:          " + comboBox14.Text;
                    worksheet.Cells[9, 1, 9, 3].Merge = true;
                }

                for (i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    worksheet.Cells[rowCount, 1].Value = Grid.Rows[i].Cells[1].Value;
                    worksheet.Cells[rowCount, 2].Value = Grid.Rows[i].Cells[2].Value;
                    worksheet.Cells[rowCount, 3].Value = Grid.Rows[i].Cells[3].Value;
                    worksheet.Cells[rowCount, 4].Value = Grid.Rows[i].Cells[4].Value;
                    worksheet.Cells[rowCount, 5].Value = Grid.Rows[i].Cells[5].Value;
                    worksheet.Cells[rowCount, 6].Value = "$ " + Grid.Rows[i].Cells[19].Value;
                    worksheet.Cells[rowCount, 7].Value = "$ " + Grid.Rows[i].Cells[20].Value;
                    worksheet.Cells[rowCount, 8].Value = "$ " + Grid.Rows[i].Cells[21].Value;
                    worksheet.Cells[rowCount, 9].Value = "$ " + Grid.Rows[i].Cells[22].Value;
                    worksheet.Cells[rowCount, 10].Value = Grid.Rows[i].Cells[23].Value;
                    worksheet.Cells[rowCount, 11].Value = Grid.Rows[i].Cells[24].Value;
                    worksheet.Cells[rowCount, 12].Value = Grid.Rows[i].Cells[25].Value;

                    totalQty = totalQty + Convert.ToInt32(Grid.Rows[i].Cells[5].Value);
                    totalAmountPriceRW = totalAmountPriceRW + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    totalAmountPriceAIR = totalAmountPriceAIR + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    rowCount++;
                }
                worksheet.Cells[rowCount, 3].Value = "Total:";
                worksheet.Cells[rowCount, 5].Value = totalQty;
                worksheet.Cells[rowCount, 8].Value = "$ " + totalAmountPriceRW;
                worksheet.Cells[rowCount, 9].Value = "$ " + totalAmountPriceAIR;
                using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 12])
                {
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 123, 129));
                }
                using (ExcelRange rng = worksheet.Cells[13, 1, rowCount, 12])
                {
                    rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                using (ExcelRange rng = worksheet.Cells[14, 1, rowCount, 3])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
                using (ExcelRange rng = worksheet.Cells[14, 4, rowCount, 12])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }

            }
            if (cBox.SelectedIndex == 1)
            {
                worksheet.Row(13).Height = 75;
                worksheet.Cells[13, 6].Value = "Unit price in USD by AIR, " + comboBox17.Text;
                worksheet.Cells[13, 7].Value = "Amount price in USD by AIR, " + comboBox17.Text;
                worksheet.Cells[13, 8].Value = "PRODUCTION TIME";
                worksheet.Cells[13, 9].Value = "Delivery by AIR";

                if (textBox15.Text.Length > 0)
                {
                    worksheet.Cells[9, 7].Value = "CP:";
                    worksheet.Cells[9, 8].Value = textBox15.Text + "-1";
                    worksheet.Cells[9, 9].Value = dateTimePicker10.Value.ToShortDateString();
                }
                if (textBox30.Text.Length > 0)
                {
                    worksheet.Cells[11, 1].Value = "ORDER:     " + textBox30.Text + "     DATE:   " + dateTimePicker15.Value.ToShortDateString();
                    worksheet.Cells[11, 1, 11, 3].Merge = true;
                }
                if (textBox35.Text.Length > 0)
                {
                    worksheet.Cells[10, 1].Value = "PROJECT:     " + textBox35.Text;
                    worksheet.Cells[10, 1, 10, 3].Merge = true;
                }
                if (comboBox14.Text.Length > 0)
                {
                    worksheet.Cells[9, 1].Value = "TO:          " + comboBox14.Text;
                    worksheet.Cells[9, 1, 9, 3].Merge = true;
                }

                for (i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    worksheet.Cells[rowCount, 1].Value = Grid.Rows[i].Cells[1].Value;
                    worksheet.Cells[rowCount, 2].Value = Grid.Rows[i].Cells[2].Value;
                    worksheet.Cells[rowCount, 3].Value = Grid.Rows[i].Cells[3].Value;
                    worksheet.Cells[rowCount, 4].Value = Grid.Rows[i].Cells[4].Value;
                    worksheet.Cells[rowCount, 5].Value = Grid.Rows[i].Cells[5].Value;
                    worksheet.Cells[rowCount, 6].Value = "$ " + Grid.Rows[i].Cells[20].Value;
                    worksheet.Cells[rowCount, 7].Value = "$ " + Grid.Rows[i].Cells[22].Value;
                    worksheet.Cells[rowCount, 8].Value = Grid.Rows[i].Cells[23].Value;
                    worksheet.Cells[rowCount, 9].Value = Grid.Rows[i].Cells[24].Value;

                    totalQty = totalQty + Convert.ToInt32(Grid.Rows[i].Cells[5].Value);
                    totalAmountPriceRW = totalAmountPriceRW + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    totalAmountPriceAIR = totalAmountPriceAIR + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    rowCount++;
                }
                worksheet.Cells[rowCount, 3].Value = "Total:";
                worksheet.Cells[rowCount, 5].Value = totalQty;
                worksheet.Cells[rowCount, 7].Value = "$ " + totalAmountPriceAIR;
                using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
                {
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 123, 129));
                }
                using (ExcelRange rng = worksheet.Cells[13, 1, rowCount, 9])
                {
                    rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                using (ExcelRange rng = worksheet.Cells[14, 1, rowCount, 3])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
                using (ExcelRange rng = worksheet.Cells[14, 4, rowCount, 9])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }
            if (cBox.SelectedIndex == 2)
            {
                worksheet.Row(13).Height = 75;
                worksheet.Cells[13, 6].Value = "Unit price in USD by RW, " + comboBox16.Text;
                worksheet.Cells[13, 7].Value = "Amount price in USD by RW, " + comboBox16.Text;
                worksheet.Cells[13, 8].Value = "PRODUCTION TIME";
                worksheet.Cells[13, 9].Value = "Delivery by RW";

                if (textBox15.Text.Length > 0)
                {
                    worksheet.Cells[9, 7].Value = "CP:";
                    worksheet.Cells[9, 8].Value = textBox15.Text + "-1";
                    worksheet.Cells[9, 9].Value = dateTimePicker10.Value.ToShortDateString();
                }
                if (textBox30.Text.Length > 0)
                {
                    worksheet.Cells[11, 1].Value = "ORDER:     " + textBox30.Text + "     DATE:   " + dateTimePicker15.Value.ToShortDateString();
                    worksheet.Cells[11, 1, 11, 3].Merge = true;
                }
                if (textBox35.Text.Length > 0)
                {
                    worksheet.Cells[10, 1].Value = "PROJECT:     " + textBox35.Text;
                    worksheet.Cells[10, 1, 10, 3].Merge = true;
                }
                if (comboBox14.Text.Length > 0)
                {
                    worksheet.Cells[9, 1].Value = "TO:          " + comboBox14.Text;
                    worksheet.Cells[9, 1, 9, 3].Merge = true;
                }

                for (i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    worksheet.Cells[rowCount, 1].Value = Grid.Rows[i].Cells[1].Value;
                    worksheet.Cells[rowCount, 2].Value = Grid.Rows[i].Cells[2].Value;
                    worksheet.Cells[rowCount, 3].Value = Grid.Rows[i].Cells[3].Value;
                    worksheet.Cells[rowCount, 4].Value = Grid.Rows[i].Cells[4].Value;
                    worksheet.Cells[rowCount, 5].Value = Grid.Rows[i].Cells[5].Value;
                    worksheet.Cells[rowCount, 6].Value = "$ " + Grid.Rows[i].Cells[19].Value;
                    worksheet.Cells[rowCount, 7].Value = "$ " + Grid.Rows[i].Cells[21].Value;
                    worksheet.Cells[rowCount, 8].Value = Grid.Rows[i].Cells[23].Value;
                    worksheet.Cells[rowCount, 9].Value = Grid.Rows[i].Cells[25].Value;

                    totalQty = totalQty + Convert.ToInt32(Grid.Rows[i].Cells[5].Value);
                    totalAmountPriceRW = totalAmountPriceRW + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    totalAmountPriceAIR = totalAmountPriceAIR + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    rowCount++;
                }
                worksheet.Cells[rowCount, 3].Value = "Total:";
                worksheet.Cells[rowCount, 5].Value = totalQty;
                worksheet.Cells[rowCount, 7].Value = "$ " + totalAmountPriceRW;
                using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
                {
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 123, 129));
                }
                using (ExcelRange rng = worksheet.Cells[13, 1, rowCount, 9])
                {
                    rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                using (ExcelRange rng = worksheet.Cells[14, 1, rowCount, 3])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
                using (ExcelRange rng = worksheet.Cells[14, 4, rowCount, 9])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }


            //using (System.Drawing.Image sign = System.Drawing.Image.FromFile(app_dir_temp + "Sign.png"))
            //{
            //    var excelImage = worksheet.Drawings.AddPicture("Sign", sign);
            //    excelImage.SetPosition(rowCount + 1, 15, 5, 0);
            //    excelImage.SetSize(14);
            //}
            using (System.Drawing.Image stamp = System.Drawing.Image.FromFile(app_dir_temp + "Stamp.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Stamp", stamp);
                excelImage.SetPosition(rowCount + 1, 0, 7, 15);
                excelImage.SetSize(20);
            }
            worksheet.Cells[rowCount + 4, 4, rowCount + 4, 9].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[rowCount + 5, 4].Value = " Authorized and signed by Anna Li";
            worksheet.Cells[rowCount + 5, 4, rowCount + 5, 9].Merge = true;
            worksheet.Cells[rowCount + 6, 4].Value = " by Power of Attorney on behalf of ";
            worksheet.Cells[rowCount + 6, 4, rowCount + 6, 9].Merge = true;
            worksheet.Cells[rowCount + 7, 4].Value = " Drone Zone General Manager  ";
            worksheet.Cells[rowCount + 7, 4, rowCount + 7, 9].Merge = true;
            worksheet.Cells[rowCount + 8, 4].Value = " dronezone.anna@gmail.com  ";
            worksheet.Cells[rowCount + 8, 4].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 4, rowCount + 8, 9].Merge = true;


            worksheet.Cells[rowCount + 7, 1].Value = " Customer care  ";
            worksheet.Cells[rowCount + 7, 1, rowCount + 7, 3].Merge = true;
            worksheet.Cells[rowCount + 8, 1].Value = " dronezone.sk@gmail.com  ";
            worksheet.Cells[rowCount + 8, 1].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 1, rowCount + 8, 3].Merge = true;

            //using (System.Drawing.Image header = System.Drawing.Image.FromFile(app_dir_temp + "cp_header.png"))
            //{
            //    var excelImage = worksheet.Drawings.AddPicture("buttom", header);
            //    excelImage.SetPosition(rowCount+8, 0, 0, 0);
            //    if (cBox.SelectedIndex == 0)
            //    {
            //        excelImage.SetSize(825, 40);
            //    }
            //    else
            //    {
            //        excelImage.SetSize(620, 40);
            //    }
            //}
            using (ExcelRange rng = worksheet.Cells[13, 1, 13, worksheet.Dimension.Columns])
            {
                rng.Style.Font.Size = 6;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            using (ExcelRange rng = worksheet.Cells[9, 1, 12, worksheet.Dimension.Columns])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            }
            using (ExcelRange rng = worksheet.Cells[9, 1, rowCount + 10, worksheet.Dimension.Columns])
            {
                rng.Style.WrapText = true;
                rng.Style.Font.Size = 6;
                rng.Style.Font.Name = "Times New Roman";
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;
            saveFileDialog.FileName = comboBox14.Text + "_" + textBox15.Text + "-1";

            if (type == 0)
            {
                // savedialog
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excel.SaveAs(new FileInfo(saveFileDialog.FileName));
                }
            }
            else
            {
                // save temp
                //Random rnd = new Random();
                //int val = rnd.Next(9999, 99999);
                string folder = contextMenuStrip7.Items[1].Text.Substring(7);
                excel.SaveAs(new FileInfo(folder + "\\" + comboBox14.Text + "_" + textBox15.Text + "-1.xlsx"));
                System.Diagnostics.Process.Start(folder + "\\"  + comboBox14.Text + "_" + textBox15.Text + "-1.xlsx");
            }
        }

        public void GenerateCP2(DataGridView Grid, string d_t, int type)
        {
            string s;
            s = textBox21.Text;
            s = s.Trim();
            textBox21.Text = s;
            //
            s = textBox22.Text;
            s = s.Trim();
            textBox22.Text = s;
            //
            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");
            worksheet.DefaultColWidth = 10;
            worksheet.Column(1).Width = 4;
            worksheet.Column(2).Width = 9;
            worksheet.Column(3).Width = 27;
            worksheet.Column(4).Width = 8;
            worksheet.Column(5).Width = 8;
            worksheet.Column(6).Width = 8;
            worksheet.Column(7).Width = 8;
            worksheet.Column(8).Width = 10;
            worksheet.Column(9).Width = 10;

            //using (System.Drawing.Image header = System.Drawing.Image.FromFile(app_dir_temp + "cp_header.png"))
            //{
            //    var excelImage = worksheet.Drawings.AddPicture("Header", header);
            //    excelImage.SetPosition(0, 0, 0, 0);
            //    if (d_t == "AIR & RW")
            //    {
            //        excelImage.SetSize(825, 40);
            //    }
            //    else
            //    {
            //        excelImage.SetSize(620, 40);
            //    }

            //}

            using (System.Drawing.Image logo = System.Drawing.Image.FromFile(app_dir_temp + "Logo.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("My Logo", logo);
                if (d_t == "AIR & RW")
                {
                    excelImage.SetPosition(0, 0, 2, 280);
                }
                else
                {
                    excelImage.SetPosition(0, 0, 2, 130);
                }
                excelImage.SetSize(25);
            }
            int selected = dataGridView10.CurrentCell.RowIndex;
            string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
            string company = dataGridView10.Rows[selected].Cells[2].Value.ToString();
            string project = dataGridView10.Rows[selected].Cells[3].Value.ToString();

            int selected_2 = dataGridView11.CurrentCell.RowIndex;
            string cp_version = dataGridView11.Rows[selected_2].Cells[0].Value.ToString();
            string cp_date = dataGridView11.Rows[selected_2].Cells[1].Value.ToString();
            string order_id = dataGridView11.Rows[selected_2].Cells[3].Value.ToString();
            string order_date = dataGridView11.Rows[selected_2].Cells[4].Value.ToString();
            string ct = dataGridView11.Rows[selected_2].Cells[5].Value.ToString();
            string ct_date = dataGridView11.Rows[selected_2].Cells[6].Value.ToString();
            string po = dataGridView11.Rows[selected_2].Cells[7].Value.ToString();
            string po_date = dataGridView11.Rows[selected_2].Cells[8].Value.ToString();


            int rowCount = 16;
            int totalQty = 0;
            double totalAmountPriceRW = 0;
            double totalAmountPriceAIR = 0;
            int i = 0;

            worksheet.Cells[15, 1].Value = "#";
            worksheet.Cells[15, 2].Value = "Item code";
            worksheet.Cells[15, 3].Value = "Description of goods";
            worksheet.Cells[15, 4].Value = "HS code";
            worksheet.Cells[15, 5].Value = "Qty";
            if (d_t == "AIR & RW")
            {
                worksheet.Row(15).Height = 75;
                worksheet.Cells[15, 6].Value = "Unit price in USD by RW, " + textBox22.Text;
                worksheet.Cells[15, 7].Value = "Unit price in USD by AIR, " + textBox21.Text;
                worksheet.Cells[15, 8].Value = "Amount price in USD by RW, " + textBox22.Text;
                worksheet.Cells[15, 9].Value = "Amount price in USD by AIR, " + textBox21.Text;
                worksheet.Cells[15, 10].Value = "PRODUCTION TIME";
                worksheet.Cells[15, 11].Value = "Delivery by AIR";
                worksheet.Cells[15, 12].Value = "Delivery by RW";

                if (cp_id.Length > 0 && cp_date.Length > 0)
                {
                    worksheet.Cells[9, 10].Value = "CP:";
                    worksheet.Cells[9, 11].Value = cp_id + "-" + cp_version;
                    worksheet.Cells[9, 12].Value = cp_date;
                }
                if (order_id.Length > 0 && order_date.Length > 0)
                {
                    worksheet.Cells[11, 1].Value = "ORDER:      " + order_id + "     DATE:   " + order_date;
                    worksheet.Cells[11, 1, 11, 3].Merge = true;
                }
                if (ct.Length > 0 && ct_date.Length > 0)
                {
                    worksheet.Cells[10, 9].Value = "CT:";
                    worksheet.Cells[10, 10].Value = ct;
                    worksheet.Cells[10, 10, 10, 11].Merge = true;
                    worksheet.Cells[10, 12].Value = ct_date;
                }
                if (po.Length > 0 && po_date.Length > 0)
                {
                    worksheet.Cells[11, 9].Value = "PO:";
                    worksheet.Cells[11, 10].Value = po;
                    worksheet.Cells[11, 10, 11, 11].Merge = true;
                    worksheet.Cells[11, 12].Value = po_date;
                }
                if (project.Length > 0)
                {
                    worksheet.Cells[10, 1].Value = "PROJECT:    " + project;
                    worksheet.Cells[10, 1, 10, 3].Merge = true;
                }
                if (company.Length > 0)
                {
                    worksheet.Cells[9, 1].Value = "TO:             " + company;
                    worksheet.Cells[9, 1, 9, 3].Merge = true;
                }

                for (i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    worksheet.Cells[rowCount, 1].Value = Grid.Rows[i].Cells[1].Value;
                    worksheet.Cells[rowCount, 2].Value = Grid.Rows[i].Cells[2].Value;
                    worksheet.Cells[rowCount, 3].Value = Grid.Rows[i].Cells[3].Value;
                    worksheet.Cells[rowCount, 4].Value = Grid.Rows[i].Cells[4].Value;
                    worksheet.Cells[rowCount, 5].Value = Grid.Rows[i].Cells[5].Value;
                    worksheet.Cells[rowCount, 6].Value = "$ " + Grid.Rows[i].Cells[19].Value;
                    worksheet.Cells[rowCount, 7].Value = "$ " + Grid.Rows[i].Cells[20].Value;
                    worksheet.Cells[rowCount, 8].Value = "$ " + Grid.Rows[i].Cells[21].Value;
                    worksheet.Cells[rowCount, 9].Value = "$ " + Grid.Rows[i].Cells[22].Value;
                    worksheet.Cells[rowCount, 10].Value = Grid.Rows[i].Cells[23].Value;
                    worksheet.Cells[rowCount, 11].Value = Grid.Rows[i].Cells[24].Value;
                    worksheet.Cells[rowCount, 12].Value = Grid.Rows[i].Cells[25].Value;

                    totalQty = totalQty + Convert.ToInt32(Grid.Rows[i].Cells[5].Value);
                    totalAmountPriceRW = totalAmountPriceRW + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    totalAmountPriceAIR = totalAmountPriceAIR + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    rowCount++;
                }
                worksheet.Cells[rowCount, 3].Value = "Total:";
                worksheet.Cells[rowCount, 5].Value = totalQty;
                worksheet.Cells[rowCount, 8].Value = "$ " + totalAmountPriceRW;
                worksheet.Cells[rowCount, 9].Value = "$ " + totalAmountPriceAIR;
                using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 12])
                {
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 123, 129));
                }
                using (ExcelRange rng = worksheet.Cells[15, 1, rowCount, 12])
                {
                    rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                using (ExcelRange rng = worksheet.Cells[14, 1, rowCount, 3])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
                using (ExcelRange rng = worksheet.Cells[14, 4, rowCount, 12])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }
            if (d_t == "AIR")
            {
                worksheet.Row(15).Height = 75;
                worksheet.Cells[15, 6].Value = "Unit price in USD by AIR, " + textBox21.Text;
                worksheet.Cells[15, 7].Value = "Amount price in USD by AIR, " + textBox21.Text;
                worksheet.Cells[15, 8].Value = "PRODUCTION TIME";
                worksheet.Cells[15, 9].Value = "Delivery by AIR";

                if (cp_id.Length > 0 && cp_date.Length > 0)
                {
                    worksheet.Cells[9, 7].Value = "CP:";
                    worksheet.Cells[9, 8].Value = cp_id + "-" + cp_version;
                    worksheet.Cells[9, 9].Value = cp_date;
                }
                if (order_id.Length > 0 && order_date.Length > 0)
                {
                    worksheet.Cells[11, 1].Value = "ORDER:      " + order_id + "     DATE:   " + order_date;
                    worksheet.Cells[11, 1, 11, 3].Merge = true;
                }
                if (ct.Length > 0 && ct_date.Length > 0)
                {
                    worksheet.Cells[10, 6].Value = "CT:";
                    worksheet.Cells[10, 7].Value = ct;
                    worksheet.Cells[10, 7, 10, 8].Merge = true;
                    worksheet.Cells[10, 9].Value = ct_date;
                }
                if (po.Length > 0 && po_date.Length > 0)
                {
                    worksheet.Cells[11, 6].Value = "PO:";
                    worksheet.Cells[11, 7].Value = po;
                    worksheet.Cells[11, 7, 11, 8].Merge = true;
                    worksheet.Cells[11, 9].Value = po_date;
                }
                if (project.Length > 0)
                {
                    worksheet.Cells[10, 1].Value = "PROJECT:    " + project;
                    worksheet.Cells[10, 1, 10, 3].Merge = true;
                }
                if (company.Length > 0)
                {
                    worksheet.Cells[9, 1].Value = "TO:             " + company;
                    worksheet.Cells[9, 1, 9, 3].Merge = true;
                }

                for (i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    worksheet.Cells[rowCount, 1].Value = Grid.Rows[i].Cells[1].Value;
                    worksheet.Cells[rowCount, 2].Value = Grid.Rows[i].Cells[2].Value;
                    worksheet.Cells[rowCount, 3].Value = Grid.Rows[i].Cells[3].Value;
                    worksheet.Cells[rowCount, 4].Value = Grid.Rows[i].Cells[4].Value;
                    worksheet.Cells[rowCount, 5].Value = Grid.Rows[i].Cells[5].Value;
                    worksheet.Cells[rowCount, 6].Value = "$ " + Grid.Rows[i].Cells[20].Value;
                    worksheet.Cells[rowCount, 7].Value = "$ " + Grid.Rows[i].Cells[22].Value;
                    worksheet.Cells[rowCount, 8].Value = Grid.Rows[i].Cells[23].Value;
                    worksheet.Cells[rowCount, 9].Value = Grid.Rows[i].Cells[24].Value;

                    totalQty = totalQty + Convert.ToInt32(Grid.Rows[i].Cells[5].Value);
                    totalAmountPriceRW = totalAmountPriceRW + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    totalAmountPriceAIR = totalAmountPriceAIR + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    rowCount++;
                }
                worksheet.Cells[rowCount, 3].Value = "Total:";
                worksheet.Cells[rowCount, 5].Value = totalQty;
                worksheet.Cells[rowCount, 7].Value = "$ " + totalAmountPriceAIR;
                using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
                {
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 123, 129));
                }
                using (ExcelRange rng = worksheet.Cells[15, 1, rowCount, 9])
                {
                    rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                using (ExcelRange rng = worksheet.Cells[14, 1, rowCount, 3])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
                using (ExcelRange rng = worksheet.Cells[14, 4, rowCount, 9])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }
            if (d_t == "RW")
            {
                worksheet.Row(15).Height = 75;
                worksheet.Cells[15, 6].Value = "Unit price in USD by RW, " + textBox22.Text;
                worksheet.Cells[15, 7].Value = "Amount price in USD by RW, " + textBox22.Text;
                worksheet.Cells[15, 8].Value = "PRODUCTION TIME";
                worksheet.Cells[15, 9].Value = "Delivery by RW";

                if (cp_id.Length > 0 && cp_date.Length > 0)
                {
                    worksheet.Cells[9, 7].Value = "CP:";
                    worksheet.Cells[9, 8].Value = cp_id + "-" + cp_version;
                    worksheet.Cells[9, 9].Value = cp_date;
                }
                if (order_id.Length > 0 && order_date.Length > 0)
                {
                    worksheet.Cells[11, 1].Value = "ORDER:      " + order_id + "     DATE:   " + order_date;
                    worksheet.Cells[11, 1, 11, 3].Merge = true;
                }
                if (ct.Length > 0 && ct_date.Length > 0)
                {
                    worksheet.Cells[10, 6].Value = "CT:";
                    worksheet.Cells[10, 7].Value = ct;
                    worksheet.Cells[10, 7, 10, 8].Merge = true;
                    worksheet.Cells[10, 9].Value = ct_date;
                }
                if (po.Length > 0 && po_date.Length > 0)
                {
                    worksheet.Cells[11, 6].Value = "PO:";
                    worksheet.Cells[11, 7].Value = po;
                    worksheet.Cells[11, 7, 11, 8].Merge = true;
                    worksheet.Cells[11, 9].Value = po_date;
                }
                if (project.Length > 0)
                {
                    worksheet.Cells[10, 1].Value = "PROJECT:    " + project;
                    worksheet.Cells[10, 1, 10, 3].Merge = true;
                }
                if (company.Length > 0)
                {
                    worksheet.Cells[9, 1].Value = "TO:             " + company;
                    worksheet.Cells[9, 1, 9, 3].Merge = true;
                }

                for (i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    worksheet.Cells[rowCount, 1].Value = Grid.Rows[i].Cells[1].Value;
                    worksheet.Cells[rowCount, 2].Value = Grid.Rows[i].Cells[2].Value;
                    worksheet.Cells[rowCount, 3].Value = Grid.Rows[i].Cells[3].Value;
                    worksheet.Cells[rowCount, 4].Value = Grid.Rows[i].Cells[4].Value;
                    worksheet.Cells[rowCount, 5].Value = Grid.Rows[i].Cells[5].Value;
                    worksheet.Cells[rowCount, 6].Value = "$ " + Grid.Rows[i].Cells[19].Value;
                    worksheet.Cells[rowCount, 7].Value = "$ " + Grid.Rows[i].Cells[21].Value;
                    worksheet.Cells[rowCount, 8].Value = Grid.Rows[i].Cells[23].Value;
                    worksheet.Cells[rowCount, 9].Value = Grid.Rows[i].Cells[25].Value;

                    totalQty = totalQty + Convert.ToInt32(Grid.Rows[i].Cells[5].Value);
                    totalAmountPriceRW = totalAmountPriceRW + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    totalAmountPriceAIR = totalAmountPriceAIR + Math.Round(Convert.ToDouble(Grid.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    rowCount++;
                }
                worksheet.Cells[rowCount, 3].Value = "Total:";
                worksheet.Cells[rowCount, 5].Value = totalQty;
                worksheet.Cells[rowCount, 7].Value = "$ " + totalAmountPriceRW;
                using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
                {
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 123, 129));
                }
                using (ExcelRange rng = worksheet.Cells[15, 1, rowCount, 9])
                {
                    rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                using (ExcelRange rng = worksheet.Cells[14, 1, rowCount, 3])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
                using (ExcelRange rng = worksheet.Cells[14, 4, rowCount, 9])
                {
                    rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }
            }

            //using (System.Drawing.Image sign = System.Drawing.Image.FromFile(app_dir_temp + "Sign.png"))
            //{
            //    var excelImage = worksheet.Drawings.AddPicture("Sign", sign);
            //    excelImage.SetPosition(rowCount + 1, 15, 5, 0);
            //    excelImage.SetSize(14);
            //}
            using (System.Drawing.Image stamp = System.Drawing.Image.FromFile(app_dir_temp + "Stamp.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Stamp", stamp);
                excelImage.SetPosition(rowCount + 1, 0, 7, 15);
                excelImage.SetSize(20);
            }
            worksheet.Cells[rowCount + 4, 4, rowCount + 4, 9].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[rowCount + 5, 4].Value = " Authorized and signed by Anna Li";
            worksheet.Cells[rowCount + 5, 4, rowCount + 5, 9].Merge = true;
            worksheet.Cells[rowCount + 6, 4].Value = " by Power of Attorney on behalf of ";
            worksheet.Cells[rowCount + 6, 4, rowCount + 6, 9].Merge = true;
            worksheet.Cells[rowCount + 7, 4].Value = " Drone Zone General Manager  ";
            worksheet.Cells[rowCount + 7, 4, rowCount + 7, 9].Merge = true;
            worksheet.Cells[rowCount + 8, 4].Value = " dronezone.anna@gmail.com  ";
            worksheet.Cells[rowCount + 8, 4].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 4, rowCount + 8, 9].Merge = true;


            worksheet.Cells[rowCount + 7, 1].Value = " Customer care  ";
            worksheet.Cells[rowCount + 7, 1, rowCount + 7, 3].Merge = true;
            worksheet.Cells[rowCount + 8, 1].Value = " dronezone.sk@gmail.com  ";
            worksheet.Cells[rowCount + 8, 1].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 1, rowCount + 8, 3].Merge = true;

            //using (System.Drawing.Image header = System.Drawing.Image.FromFile(app_dir_temp + "cp_header.png"))
            //{
            //    var excelImage = worksheet.Drawings.AddPicture("buttom", header);
            //    excelImage.SetPosition(rowCount + 8, 0, 0, 0);
            //    if (d_t == "AIR & RW")
            //    {
            //        excelImage.SetSize(825, 40);
            //    }
            //    else
            //    {
            //        excelImage.SetSize(620, 40);
            //    }
            //}

            using (ExcelRange rng = worksheet.Cells[15, 1, 15, worksheet.Dimension.Columns])
            {
                rng.Style.Font.Size = 6;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            using (ExcelRange rng = worksheet.Cells[9, 1, 12, worksheet.Dimension.Columns])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            }
            using (ExcelRange rng = worksheet.Cells[9, 1, rowCount + 10, worksheet.Dimension.Columns])
            {
                rng.Style.WrapText = true;
                rng.Style.Font.Size = 6;
                rng.Style.Font.Name = "Times New Roman";
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;
            saveFileDialog.FileName = company + "_" + cp_id + "-" + cp_version;

            if (type == 0)
            {
                // savedialog
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excel.SaveAs(new FileInfo(saveFileDialog.FileName));
                }
            }
            else
            {
                // save temp
                //Random rnd = new Random();
                //int val = rnd.Next(9999, 99999);
                int vc = dataGridView11.RowCount;
                //
                string folder = contextMenuStrip7.Items[1].Text.Substring(7);
                excel.SaveAs(new FileInfo(folder + "\\" + company + "_" + cp_id + "-" + (vc) + ".xlsx"));
                System.Diagnostics.Process.Start(folder + "\\" + company + "_" + cp_id + "-" + (vc) + ".xlsx");
            }
        }

        CheckBox checkAllCheckBox2;
        void checkBox2_Changed(object sender, EventArgs e)
        {
            for (int j = 0; j < this.dataGridView12.RowCount; j++)
            {
                this.dataGridView12[0, j].Value = this.checkAllCheckBox2.Checked;
            }
            this.dataGridView12.EndEdit();
        }

        public void InitGridView12()
        {
            dataGridView12.Rows.Clear();
            dataGridView12.Refresh();
            dataGridView12.Controls.Clear();
            dataGridView12.ColumnCount = 0;

            dataGridView12.RowCount = 0;
            dataGridView12.ColumnCount = 25;

            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "  ";
            checkColumn.Width = 25;
            checkColumn.ReadOnly = false;
            dataGridView12.Columns.Insert(0, checkColumn);

            checkAllCheckBox2 = new CheckBox();
            checkAllCheckBox2.Size = new Size(15, 15);
            Rectangle rect = dataGridView12.GetCellDisplayRectangle(0, -1, true);
            rect.Y = dataGridView12.ColumnHeadersHeight;
            rect.X = checkColumn.Width + dataGridView12.RowHeadersWidth / 2;
            checkAllCheckBox2.Location = rect.Location;
            checkAllCheckBox2.CheckedChanged += new EventHandler(checkBox2_Changed);
            dataGridView12.Controls.Add(checkAllCheckBox2);

            dataGridView12.Columns[1].HeaderText = "#";
            dataGridView12.Columns[2].HeaderText = "Код продукта";
            dataGridView12.Columns[3].HeaderText = "Название продукта";
            dataGridView12.Columns[4].HeaderText = "HS код";
            dataGridView12.Columns[5].HeaderText = "Кол-во";
            dataGridView12.Columns[6].HeaderText = "Цена за ед.";
            dataGridView12.Columns[7].HeaderText = "Общая цена";
            dataGridView12.Columns[8].HeaderText = "% ";
            dataGridView12.Columns[9].HeaderText = "Доход с ед.";
            dataGridView12.Columns[10].HeaderText = "Сумма дохода";
            dataGridView12.Columns[11].HeaderText = "Оригинал. цена за ед. + %";
            dataGridView12.Columns[12].HeaderText = "Сумма оригинал. цена за ед. + %";
            dataGridView12.Columns[13].HeaderText = "Вес 1 ед.";
            dataGridView12.Columns[14].HeaderText = "Общий вес";
            dataGridView12.Columns[15].HeaderText = "Стоимость доставки 1 кг - RW";
            dataGridView12.Columns[16].HeaderText = "Общая стоимость доставки - RW";
            dataGridView12.Columns[17].HeaderText = "Стоимость доставки 1 кг - AIR"; // Стоимость доставки 1 кг - AIR, DAP Tashkent Airport, Islam Karimov
            dataGridView12.Columns[18].HeaderText = "Общая стоимость доставки - AIR";
            dataGridView12.Columns[19].HeaderText = "Цена за ед. (USD) - RW"; // CPT Tashkent,Uzbekistan,Sergeli stantion
            dataGridView12.Columns[20].HeaderText = "Цена за ед. (USD) - AIR";
            dataGridView12.Columns[21].HeaderText = "Общая стоимость (USD) - RW";
            dataGridView12.Columns[22].HeaderText = "Общая стоимость (USD) - AIR";
            dataGridView12.Columns[23].HeaderText = "Срок производства";
            dataGridView12.Columns[24].HeaderText = "Срок доставки  AIR";
            dataGridView12.Columns[25].HeaderText = "Срок доставки  RW";

            dataGridView12.Columns[1].ReadOnly = true;
            dataGridView12.Columns[7].ReadOnly = true;
            dataGridView12.Columns[9].ReadOnly = true;
            dataGridView12.Columns[10].ReadOnly = true;
            dataGridView12.Columns[11].ReadOnly = true;
            dataGridView12.Columns[12].ReadOnly = true;
            dataGridView12.Columns[14].ReadOnly = true;
            dataGridView12.Columns[16].ReadOnly = true;
            dataGridView12.Columns[18].ReadOnly = true;
            dataGridView12.Columns[21].ReadOnly = true;
            dataGridView12.Columns[22].ReadOnly = true;

            dataGridView12.Columns[1].Width = 40;
            dataGridView12.Columns[2].Width = 85;
            dataGridView12.Columns[3].Width = 230;
            dataGridView12.Columns[4].Width = 70;
            dataGridView12.Columns[5].Width = 65;
            dataGridView12.Columns[6].Width = 65;
            dataGridView12.Columns[7].Width = 65;
            dataGridView12.Columns[8].Width = 50;
            dataGridView12.Columns[9].Width = 70;
            dataGridView12.Columns[10].Width = 70;
            dataGridView12.Columns[11].Width = 70;
            dataGridView12.Columns[12].Width = 70;
            dataGridView12.Columns[13].Width = 60;
            dataGridView12.Columns[14].Width = 60;
            dataGridView12.Columns[15].Width = 70;
            dataGridView12.Columns[16].Width = 70;
            dataGridView12.Columns[17].Width = 70;
            dataGridView12.Columns[18].Width = 70;
            dataGridView12.Columns[19].Width = 85;
            dataGridView12.Columns[20].Width = 85;
            dataGridView12.Columns[21].Width = 85;
            dataGridView12.Columns[22].Width = 85;
            dataGridView12.Columns[23].Width = 85;
            dataGridView12.Columns[24].Width = 70;
            dataGridView12.Columns[25].Width = 70;
            //
            dataGridView12.EnableHeadersVisualStyles = false;
            dataGridView12.Columns[2].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[3].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[4].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[5].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[6].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[8].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[13].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[15].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[17].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[19].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[19].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[20].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[23].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[24].HeaderCell.Style.BackColor = cl_even;
            dataGridView12.Columns[25].HeaderCell.Style.BackColor = cl_even;
           
            foreach (DataGridViewColumn col in dataGridView12.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
        }

        public class CP
        {
            public string part_code { get; set; }
            public string part_name { get; set; }
            public string hs_code { get; set; }
            public string qty { get; set; }
            public string unit_p { get; set; }
            public string amount_p { get; set; }
            public string perc { get; set; }
            public string w_item { get; set; }
            public string rw_p { get; set; }
            public string air_p { get; set; }
            public string production { get; set; }
            public string delivery_rw { get; set; }
            public string delivery_air { get; set; }
        }

        public void LoadCPGrid2(string doc)
        {
            int indexOfTotalRow = 0;
            double defaultPercentage = 71;
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (doc == "") {
                // open from dialog
                //
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo file = new FileInfo(openFileDialog.FileName);
                    ExcelPackage excel = new ExcelPackage(file);
                    var worksheet = excel.Workbook.Worksheets[1];
                    //
                    if (dataGridView12.RowCount > 0) {
                        dataGridView12.Rows.RemoveAt(dataGridView12.RowCount - 1);
                    }
                    int row_c = dataGridView12.RowCount;
                    dataGridView12.RowCount = dataGridView12.RowCount + (worksheet.Dimension.Rows - 1);
                    indexOfTotalRow = dataGridView12.RowCount;
                    //if (row_c < 0) { dataGridView12.RowCount = dataGridView12.RowCount + (worksheet.Dimension.Rows - 1); }
                    //
                    if (textBox24.Text == "") { textBox24.Text = "0"; }
                    if (textBox23.Text == "") { textBox23.Text = "7"; }
                    //
                    for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                    {
                        dataGridView12.Rows[k + row_c].Cells[1].Value = k + row_c + 1;
                        //dataGridView12.Rows[k + row_c].Cells[1].Value = worksheet.Cells[k + 2, 1].Value;
                        var ws5 = worksheet.Cells[k + 2, 5].Value;
                        var ws6 = worksheet.Cells[k + 2, 6].Value;
                        if (ws5 == null) { ws5 = 0; }
                        if (ws6 == null) { ws6 = 0; }
                        dataGridView12.Rows[k + row_c].Cells[2].Value = worksheet.Cells[k + 2, 2].Value;
                        dataGridView12.Rows[k + row_c].Cells[3].Value = worksheet.Cells[k + 2, 3].Value;
                        dataGridView12.Rows[k + row_c].Cells[4].Value = worksheet.Cells[k + 2, 4].Value;
                        dataGridView12.Rows[k + row_c].Cells[5].Value = ws5;
                        dataGridView12.Rows[k + row_c].Cells[6].Value = ws6;
                        //dataGridView12.Rows[k + row_c].Cells[7].Value = worksheet.Cells[k + 2, 7].Value;
                        dataGridView12.Rows[k + row_c].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[8].Value = defaultPercentage;
                        dataGridView12.Rows[k + row_c].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        var ws8 = worksheet.Cells[k + 2, 8].Value;
                        if (ws8 == null) { ws8 = 0; }
                        dataGridView12.Rows[k + row_c].Cells[13].Value = ws8;
                        dataGridView12.Rows[k + row_c].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[15].Value = textBox24.Text;
                        dataGridView12.Rows[k + row_c].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[17].Value = textBox23.Text;
                        dataGridView12.Rows[k + row_c].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[k + row_c].Cells[23].Value = "2-3 week";
                        dataGridView12.Rows[k + row_c].Cells[24].Value = "2-3 days";
                        dataGridView12.Rows[k + row_c].Cells[25].Value = "3-4 week";
                        //
                        if (k % 2 != 0)
                        {
                            dataGridView12.Rows[k + row_c].Cells[0].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[1].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[2].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[3].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[4].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[5].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[6].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[7].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                            dataGridView12.Rows[k + row_c].Cells[9].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[10].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[11].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[12].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[13].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[14].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[15].Style.BackColor = Color.SandyBrown; //
                            dataGridView12.Rows[k + row_c].Cells[16].Style.BackColor = Color.SandyBrown; //
                            dataGridView12.Rows[k + row_c].Cells[17].Style.BackColor = Color.SkyBlue;
                            dataGridView12.Rows[k + row_c].Cells[18].Style.BackColor = Color.SkyBlue;
                            dataGridView12.Rows[k + row_c].Cells[19].Style.BackColor = Color.SandyBrown;
                            dataGridView12.Rows[k + row_c].Cells[20].Style.BackColor = Color.SkyBlue;
                            dataGridView12.Rows[k + row_c].Cells[21].Style.BackColor = Color.SandyBrown; //
                            dataGridView12.Rows[k + row_c].Cells[22].Style.BackColor = Color.SkyBlue; //
                            dataGridView12.Rows[k + row_c].Cells[23].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[24].Style.BackColor = cl_even;
                            dataGridView12.Rows[k + row_c].Cells[25].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView12.Rows[k + row_c].Cells[0].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[1].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[2].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[3].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[4].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[5].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[6].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[7].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                            dataGridView12.Rows[k + row_c].Cells[9].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[10].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[11].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[12].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[13].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[14].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[15].Style.BackColor = Color.SandyBrown; //
                            dataGridView12.Rows[k + row_c].Cells[16].Style.BackColor = Color.SandyBrown; //
                            dataGridView12.Rows[k + row_c].Cells[17].Style.BackColor = Color.LightSkyBlue;
                            dataGridView12.Rows[k + row_c].Cells[18].Style.BackColor = Color.LightSkyBlue;
                            dataGridView12.Rows[k + row_c].Cells[19].Style.BackColor = Color.SandyBrown;
                            dataGridView12.Rows[k + row_c].Cells[20].Style.BackColor = Color.LightSkyBlue;
                            dataGridView12.Rows[k + row_c].Cells[21].Style.BackColor = Color.SandyBrown; //
                            dataGridView12.Rows[k + row_c].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                            dataGridView12.Rows[k + row_c].Cells[23].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[24].Style.BackColor = cl_odd;
                            dataGridView12.Rows[k + row_c].Cells[25].Style.BackColor = cl_odd;
                            // SeaGreen MediumSeaGreen DarkSeaGreen
                            // AliceBlue CornflowerBlue AliceBlue
                        }
                    }
                    dataGridView12.RowCount = dataGridView12.RowCount + 1;
                    //
                    dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
                    dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;

                    dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
                    dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
                    dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
                    dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
                    dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
                    dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
                    dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;


                    double counter = 0;
                    for (int k = 0; k < indexOfTotalRow; k++)
                    {
                        dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                        dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        counter++;
                    }
                    dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                }
            } else {
                // open from template
                //
                FileInfo file = new FileInfo(doc);
                ExcelPackage excel = new ExcelPackage(file);
                var worksheet = excel.Workbook.Worksheets[1];
                //
                int row_c = dataGridView12.RowCount - 1;

                if (row_c < 0)
                {
                    row_c = 0;
                }
                if (dataGridView12.RowCount == 0){
                    dataGridView12.RowCount = dataGridView12.RowCount + (worksheet.Dimension.Rows);
                }
                else {
                    dataGridView12.RowCount = dataGridView12.RowCount + (worksheet.Dimension.Rows - 1);
                }
                //dataGridView12.RowCount = dataGridView12.RowCount + (worksheet.Dimension.Rows);
                //dataGridView12.RowCount = worksheet.Dimension.Rows;

                indexOfTotalRow = dataGridView12.RowCount - 1;
                //if (row_c < 0) { dataGridView12.RowCount = dataGridView12.RowCount + (worksheet.Dimension.Rows - 1); }
                //
                if (textBox24.Text == "") { textBox24.Text = "0"; }
                if (textBox23.Text == "") { textBox23.Text = "7"; }
                //
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView12.Rows[k + row_c].Cells[1].Value = k + row_c + 1;
                    //dataGridView12.Rows[k + row_c].Cells[1].Value = worksheet.Cells[k + 2, 1].Value;

                    var ws5 = worksheet.Cells[k + 2, 5].Value;
                    var ws6 = worksheet.Cells[k + 2, 6].Value;
                    if (ws5 == null) { ws5 = 0; }
                    if (ws6 == null) { ws6 = 0; }
                    dataGridView12.Rows[k + row_c].Cells[2].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView12.Rows[k + row_c].Cells[3].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView12.Rows[k + row_c].Cells[4].Value = worksheet.Cells[k + 2, 4].Value;
                    dataGridView12.Rows[k + row_c].Cells[5].Value = ws5;
                    dataGridView12.Rows[k + row_c].Cells[6].Value = ws6;
                    //dataGridView12.Rows[k + row_c].Cells[7].Value = worksheet.Cells[k + 2, 7].Value;
                    dataGridView12.Rows[k + row_c].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[8].Value = defaultPercentage;
                    dataGridView12.Rows[k + row_c].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    var ws8 = worksheet.Cells[k + 2, 8].Value;
                    if (ws8 == null) { ws8 = 0; }
                    dataGridView12.Rows[k + row_c].Cells[13].Value = ws8;
                    dataGridView12.Rows[k + row_c].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[15].Value = textBox24.Text;
                    dataGridView12.Rows[k + row_c].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[17].Value = textBox23.Text;
                    dataGridView12.Rows[k + row_c].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k + row_c].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[k + row_c].Cells[23].Value = "2-3 week";
                    dataGridView12.Rows[k + row_c].Cells[24].Value = "2-3 days";
                    dataGridView12.Rows[k + row_c].Cells[25].Value = "3-4 week";
                    //
                    if (k % 2 == 0)
                    {
                        dataGridView12.Rows[k + row_c].Cells[0].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[1].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[2].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[3].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[4].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[5].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[6].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[7].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                        dataGridView12.Rows[k + row_c].Cells[9].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[10].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[11].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[12].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[13].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[14].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[15].Style.BackColor = Color.SandyBrown; //
                        dataGridView12.Rows[k + row_c].Cells[16].Style.BackColor = Color.SandyBrown; //
                        dataGridView12.Rows[k + row_c].Cells[17].Style.BackColor = Color.SkyBlue;
                        dataGridView12.Rows[k + row_c].Cells[18].Style.BackColor = Color.SkyBlue;
                        dataGridView12.Rows[k + row_c].Cells[19].Style.BackColor = Color.SandyBrown;
                        dataGridView12.Rows[k + row_c].Cells[20].Style.BackColor = Color.SkyBlue;
                        dataGridView12.Rows[k + row_c].Cells[21].Style.BackColor = Color.SandyBrown; //
                        dataGridView12.Rows[k + row_c].Cells[22].Style.BackColor = Color.SkyBlue; //
                        dataGridView12.Rows[k + row_c].Cells[23].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[24].Style.BackColor = cl_even;
                        dataGridView12.Rows[k + row_c].Cells[25].Style.BackColor = cl_even;
                    }
                    else
                    {
                        dataGridView12.Rows[k + row_c].Cells[0].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[1].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[2].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[3].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[4].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[5].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[6].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[7].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                        dataGridView12.Rows[k + row_c].Cells[9].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[10].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[11].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[12].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[13].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[14].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[15].Style.BackColor = Color.SandyBrown; //
                        dataGridView12.Rows[k + row_c].Cells[16].Style.BackColor = Color.SandyBrown; //
                        dataGridView12.Rows[k + row_c].Cells[17].Style.BackColor = Color.LightSkyBlue;
                        dataGridView12.Rows[k + row_c].Cells[18].Style.BackColor = Color.LightSkyBlue;
                        dataGridView12.Rows[k + row_c].Cells[19].Style.BackColor = Color.SandyBrown;
                        dataGridView12.Rows[k + row_c].Cells[20].Style.BackColor = Color.LightSkyBlue;
                        dataGridView12.Rows[k + row_c].Cells[21].Style.BackColor = Color.SandyBrown; //
                        dataGridView12.Rows[k + row_c].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                        dataGridView12.Rows[k + row_c].Cells[23].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[24].Style.BackColor = cl_odd;
                        dataGridView12.Rows[k + row_c].Cells[25].Style.BackColor = cl_odd;
                        // SeaGreen MediumSeaGreen DarkSeaGreen
                        // AliceBlue CornflowerBlue AliceBlue
                    }
                }
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;

                dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

                double counter = 0;
                for (int k = 0; k < indexOfTotalRow; k++)
                {
                    dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                    dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    counter++;
                }
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            }
            
            dataGridView12.ClearSelection();
        }

        public void GetCPItems(string cp_id, int v)
        {
            List<CP> cp_items = new List<CP>();
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT rw, air FROM com_proposal WHERE cp_id='" + cp_id + "' and version=" + v, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            textBox22.Text = reader.GetValue(0).ToString();
                            textBox21.Text = reader.GetValue(1).ToString();
                        }
                    }
                }
                connection.Close();
            }
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT * FROM items_in_cp WHERE cp_id='" + cp_id + "' and version=" + v, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CP item = new CP();
                            item.part_code = reader.GetValue(1).ToString();
                            item.part_name = reader.GetValue(2).ToString();
                            item.hs_code = reader.GetValue(3).ToString();
                            item.qty = reader.GetValue(4).ToString();
                            item.unit_p = reader.GetValue(5).ToString();
                            item.amount_p = reader.GetValue(6).ToString();
                            item.perc = reader.GetValue(8).ToString();
                            item.w_item = reader.GetValue(9).ToString();
                            item.rw_p = reader.GetValue(10).ToString();
                            item.air_p = reader.GetValue(11).ToString();
                            item.production = reader.GetValue(12).ToString();
                            item.delivery_air = reader.GetValue(13).ToString();
                            item.delivery_rw = reader.GetValue(14).ToString();
                            cp_items.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            dataGridView12.RowCount = 0;
            dataGridView12.RowCount = cp_items.Count + 1;
            dataGridView12.RowHeadersWidth = 35;
            int indexOfTotalRow = cp_items.Count;
            for (int i = 0; i < cp_items.Count; i++)
            {
                double a = 0;
                double b = 0;
                dataGridView12.Rows[i].Cells[1].Value = i + 1;
                dataGridView12.Rows[i].Cells[2].Value = cp_items[i].part_code;
                dataGridView12.Rows[i].Cells[3].Value = cp_items[i].part_name;
                dataGridView12.Rows[i].Cells[4].Value = cp_items[i].hs_code;
                a = Math.Round(Convert.ToDouble(cp_items[i].qty.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                b = Math.Round(Convert.ToDouble(cp_items[i].unit_p.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                
                dataGridView12.Rows[i].Cells[5].Value = a;
                dataGridView12.Rows[i].Cells[6].Value = b;
                dataGridView12.Rows[i].Cells[7].Value = a * b;
                dataGridView12.Rows[i].Cells[8].Value = cp_items[i].perc;
                dataGridView12.Rows[i].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[13].Value = cp_items[i].w_item;
                dataGridView12.Rows[i].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[15].Value = cp_items[i].rw_p;
                dataGridView12.Rows[i].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[17].Value = cp_items[i].air_p;
                dataGridView12.Rows[i].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[i].Cells[23].Value = cp_items[i].production;
                dataGridView12.Rows[i].Cells[24].Value = cp_items[i].delivery_rw;
                dataGridView12.Rows[i].Cells[25].Value = cp_items[i].delivery_air;

                if (i % 2 == 0)
                {
                    dataGridView12.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[5].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[6].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[7].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                    dataGridView12.Rows[i].Cells[9].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[10].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[11].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[12].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[13].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[14].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[15].Style.BackColor = Color.SandyBrown; //
                    dataGridView12.Rows[i].Cells[16].Style.BackColor = Color.SandyBrown; //
                    dataGridView12.Rows[i].Cells[17].Style.BackColor = Color.SkyBlue;
                    dataGridView12.Rows[i].Cells[18].Style.BackColor = Color.SkyBlue;
                    dataGridView12.Rows[i].Cells[19].Style.BackColor = Color.SandyBrown;
                    dataGridView12.Rows[i].Cells[20].Style.BackColor = Color.SkyBlue;
                    dataGridView12.Rows[i].Cells[21].Style.BackColor = Color.SandyBrown; //
                    dataGridView12.Rows[i].Cells[22].Style.BackColor = Color.SkyBlue; //
                    dataGridView12.Rows[i].Cells[23].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[24].Style.BackColor = cl_even;
                    dataGridView12.Rows[i].Cells[25].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView12.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[5].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[6].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[7].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                    dataGridView12.Rows[i].Cells[9].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[10].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[11].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[12].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[13].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[14].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[15].Style.BackColor = Color.SandyBrown; //
                    dataGridView12.Rows[i].Cells[16].Style.BackColor = Color.SandyBrown; //
                    dataGridView12.Rows[i].Cells[17].Style.BackColor = Color.LightSkyBlue;
                    dataGridView12.Rows[i].Cells[18].Style.BackColor = Color.LightSkyBlue;
                    dataGridView12.Rows[i].Cells[19].Style.BackColor = Color.SandyBrown;
                    dataGridView12.Rows[i].Cells[20].Style.BackColor = Color.LightSkyBlue;
                    dataGridView12.Rows[i].Cells[21].Style.BackColor = Color.SandyBrown; //
                    dataGridView12.Rows[i].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                    dataGridView12.Rows[i].Cells[23].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[24].Style.BackColor = cl_odd;
                    dataGridView12.Rows[i].Cells[25].Style.BackColor = cl_odd;
                    // SeaGreen MediumSeaGreen DarkSeaGreen
                    // AliceBlue CornflowerBlue AliceBlue
                }
            }
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;

            dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        public void CPGridCellChange2(int row, int column)
        {
            if (column == 8)
            {
                dataGridView12.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);

            }
            if (column == 13)
            {
                dataGridView12.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 15)
            {
                dataGridView12.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 17)
            {
                dataGridView12.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 19)
            {
                dataGridView12.Rows[row].Cells[8].Value = Math.Round(((Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) - (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) - Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) / Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) * 100, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 20)
            {
                dataGridView12.Rows[row].Cells[8].Value = Math.Round(((Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) - (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) - Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) / Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) * 100, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 5)
            {
                dataGridView12.Rows[row].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 6)
            {
                dataGridView12.Rows[row].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            int indexOfTotalRow = dataGridView12.RowCount - 1;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;

            dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        public void RefreshCPGrid2(string text, int type)
        {
            if (text.Length > 0)
            {
                // Refresh %
                if (type == 0)
                {
                    for (int k = 0; k < dataGridView12.Rows.Count-1; k++)
                    {
                        if (Convert.ToBoolean(dataGridView12.Rows[k].Cells[0].Value))
                        {
                            dataGridView12.Rows[k].Cells[8].Value = Math.Round(Convert.ToDouble(textBox25.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        }
                    }
                }
                // Refresh RW
                if (type == 1)
                {
                    for (int k = 0; k < dataGridView12.Rows.Count-1; k++)
                    {
                        if (Convert.ToBoolean(dataGridView12.Rows[k].Cells[0].Value))
                        {
                            dataGridView12.Rows[k].Cells[15].Value = Math.Round(Convert.ToDouble(textBox24.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        }
                    }
                }
                // Refresh AIR
                if (type == 2)
                {
                    for (int k = 0; k < dataGridView12.Rows.Count-1; k++)
                    {
                        if (Convert.ToBoolean(dataGridView12.Rows[k].Cells[0].Value))
                        {
                            dataGridView12.Rows[k].Cells[17].Value = Math.Round(Convert.ToDouble(textBox23.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView12.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        }
                    }
                }

                int indexOfTotalRow = dataGridView12.RowCount - 1;
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;

                dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
                dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
                dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

                double counter = 0;
                for (int k = 0; k < indexOfTotalRow; k++)
                {
                    dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                    dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    counter++;
                }
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            }
        }

        public void UploadCP2(string cp_id, int v, string customer, string prj)
        {
            string cp_date = "";
            string partCode = "";
            string partName = "";
            string hsCode = "";
            string quantity = "";
            string unitPrice = "";
            string amountPrice = "";
            string percentage = "";
            string weight = "";
            string priceRw = "";
            string priceAir = "";
            string poTime = "";
            string deliveryAIR = "";
            string deliveryRW = "";
            //
            string year = Convert.ToString(dateTimePicker11.Value.Year);
            string month = Convert.ToString(dateTimePicker11.Value.Month);
            string day = Convert.ToString(dateTimePicker11.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            cp_date = day + "." + month + "." + year;
            //
            year = Convert.ToString(dateTimePicker16.Value.Year);
            month = Convert.ToString(dateTimePicker16.Value.Month);
            day = Convert.ToString(dateTimePicker16.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string order_date = day + "." + month + "." + year;
            string order_id = textBox32.Text;

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO com_proposal(cp_id, cp_date, version, delivery_type, rw, air, order_id, order_date, customer, project) VALUES(@cpid, @cpdate, @v, @delivery, @rw, @air, @order, @order_dt, @customer, @prj)";
            command1.Parameters.AddWithValue("@cpid", cp_id);
            command1.Parameters.AddWithValue("@cpdate", cp_date);
            command1.Parameters.AddWithValue("@v", v + 1);
            command1.Parameters.AddWithValue("@delivery", comboBox6.Text);
            command1.Parameters.AddWithValue("@rw", textBox22.Text);
            command1.Parameters.AddWithValue("@air", textBox21.Text);
            command1.Parameters.AddWithValue("@order", order_id);
            command1.Parameters.AddWithValue("@order_dt", order_date);
            command1.Parameters.AddWithValue("@customer", customer);
            command1.Parameters.AddWithValue("@prj", prj);
            command1.ExecuteNonQuery();
            //
            int i = 0;
            for (i = 0; i < dataGridView12.Rows.Count-1; i++)
            {
                try { partCode = dataGridView12.Rows[i].Cells[2].Value.ToString(); } catch { partCode = ""; }
                try { partName = dataGridView12.Rows[i].Cells[3].Value.ToString(); } catch { partName = ""; }
                try { hsCode = dataGridView12.Rows[i].Cells[4].Value.ToString(); } catch { hsCode = ""; }
                try { quantity = dataGridView12.Rows[i].Cells[5].Value.ToString(); } catch { quantity = ""; }
                try { unitPrice = dataGridView12.Rows[i].Cells[6].Value.ToString(); } catch { unitPrice = ""; }
                try { amountPrice = dataGridView12.Rows[i].Cells[7].Value.ToString(); } catch { amountPrice = ""; }
                try { percentage = dataGridView12.Rows[i].Cells[8].Value.ToString(); } catch { percentage = ""; }
                try { weight = dataGridView12.Rows[i].Cells[13].Value.ToString(); } catch { weight = ""; }
                try { priceRw = dataGridView12.Rows[i].Cells[15].Value.ToString(); } catch { priceRw = ""; }
                try { priceAir = dataGridView12.Rows[i].Cells[17].Value.ToString(); } catch { priceAir = ""; }
                try { poTime = dataGridView12.Rows[i].Cells[23].Value.ToString(); } catch { poTime = ""; }
                try { deliveryAIR = dataGridView12.Rows[i].Cells[24].Value.ToString(); } catch { deliveryAIR = ""; }
                try { deliveryRW = dataGridView12.Rows[i].Cells[25].Value.ToString(); } catch { deliveryRW = ""; }

                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "INSERT INTO items_in_cp(part_code, part_name, hs_code, quantity, unit_price, amount_price, cp_id, percentage, weight_of_item, price_rw, price_air, production_time, delivery_air, delivery_rw, version) VALUES(@partcode, @partname, @hscode, @quantity, @unitprice, @amountprice, @cpid, @percentage, @weight, @pricerw, @priceair, @potime, @dair, @drw, @v)";
                command.Parameters.AddWithValue("@partcode", partCode);
                command.Parameters.AddWithValue("@partname", partName);
                command.Parameters.AddWithValue("@hscode", hsCode);
                command.Parameters.AddWithValue("@quantity", quantity);
                command.Parameters.AddWithValue("@unitprice", unitPrice);
                command.Parameters.AddWithValue("@amountprice", amountPrice);
                command.Parameters.AddWithValue("@cpid", cp_id);
                command.Parameters.AddWithValue("@percentage", percentage);
                command.Parameters.AddWithValue("@weight", weight);
                command.Parameters.AddWithValue("@pricerw", priceRw);
                command.Parameters.AddWithValue("@priceair", priceAir);
                command.Parameters.AddWithValue("@potime", poTime);
                command.Parameters.AddWithValue("@dair", deliveryAIR);
                command.Parameters.AddWithValue("@drw", deliveryRW);
                command.Parameters.AddWithValue("@v", v + 1);
                command.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("Комерческое предложение сохраненно.", "Done",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
            button24.Enabled = true;
            button25.Enabled = false;
            dataGridView10.Enabled = true;
            dataGridView11.Enabled = true;
            textBox32.Clear();
        }

        // CP
        CheckBox checkAllCheckBox;
        void checkBox_Changed(object sender, EventArgs e)
        {
            for (int j = 0; j < this.dataGridView9.RowCount; j++)
            {
                this.dataGridView9[0, j].Value = this.checkAllCheckBox.Checked;
            }
            this.dataGridView9.EndEdit();
        }

        public void InitGridView9()
        {
            dataGridView9.Rows.Clear();
            dataGridView9.Refresh();
            dataGridView9.Controls.Clear();
            dataGridView9.ColumnCount = 0;

            dataGridView9.RowCount = 0;
            dataGridView9.ColumnCount = 25;

            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "  ";
            checkColumn.Width = 25;
            checkColumn.ReadOnly = false;
            dataGridView9.Columns.Insert(0, checkColumn);

            checkAllCheckBox = new CheckBox();
            checkAllCheckBox.Size = new Size(15, 15);
            Rectangle rect = dataGridView9.GetCellDisplayRectangle(0, -1, true);
            rect.Y = dataGridView9.ColumnHeadersHeight;
            rect.X = checkColumn.Width + dataGridView9.RowHeadersWidth / 2;
            checkAllCheckBox.Location = rect.Location;
            checkAllCheckBox.CheckedChanged += new EventHandler(checkBox_Changed);
            dataGridView9.Controls.Add(checkAllCheckBox);

            dataGridView9.Columns[1].HeaderText = "#";
            dataGridView9.Columns[2].HeaderText = "Код продукта";
            dataGridView9.Columns[3].HeaderText = "Название продукта";
            dataGridView9.Columns[4].HeaderText = "HS код";
            dataGridView9.Columns[5].HeaderText = "Кол-во";
            dataGridView9.Columns[6].HeaderText = "Цена за ед.";
            dataGridView9.Columns[7].HeaderText = "Общая цена";
            dataGridView9.Columns[8].HeaderText = "% ";
            dataGridView9.Columns[9].HeaderText = "Доход с ед.";
            dataGridView9.Columns[10].HeaderText = "Сумма дохода";
            dataGridView9.Columns[11].HeaderText = "Оригинал. цена за ед. + %";
            dataGridView9.Columns[12].HeaderText = "Сумма оригинал. цена за ед. + %";
            dataGridView9.Columns[13].HeaderText = "Вес 1 ед.";
            dataGridView9.Columns[14].HeaderText = "Общий вес";
            dataGridView9.Columns[15].HeaderText = "Стоимость доставки 1 кг - RW";
            dataGridView9.Columns[16].HeaderText = "Общая стоимость доставки - RW";
            dataGridView9.Columns[17].HeaderText = "Стоимость доставки 1 кг - AIR"; // Стоимость доставки 1 кг - AIR, DAP Tashkent Airport, Islam Karimov
            dataGridView9.Columns[18].HeaderText = "Общая стоимость доставки - AIR";
            dataGridView9.Columns[19].HeaderText = "Цена за ед. (USD) - RW"; // CPT Tashkent,Uzbekistan,Sergeli stantion
            dataGridView9.Columns[20].HeaderText = "Цена за ед. (USD) - AIR";
            dataGridView9.Columns[21].HeaderText = "Общая стоимость (USD) - RW";
            dataGridView9.Columns[22].HeaderText = "Общая стоимость (USD) - AIR";
            dataGridView9.Columns[23].HeaderText = "Срок производства";
            dataGridView9.Columns[24].HeaderText = "Срок доставки  AIR";
            dataGridView9.Columns[25].HeaderText = "Срок доставки  RW";

            dataGridView9.Columns[1].ReadOnly = true;
            dataGridView9.Columns[2].ReadOnly = false;
            dataGridView9.Columns[3].ReadOnly = false;
            dataGridView9.Columns[4].ReadOnly = false;
            dataGridView9.Columns[5].ReadOnly = false;
            dataGridView9.Columns[6].ReadOnly = false;
            dataGridView9.Columns[7].ReadOnly = true;
            dataGridView9.Columns[9].ReadOnly = true;
            dataGridView9.Columns[10].ReadOnly = true;
            dataGridView9.Columns[11].ReadOnly = true;
            dataGridView9.Columns[12].ReadOnly = true;
            dataGridView9.Columns[14].ReadOnly = true;
            dataGridView9.Columns[16].ReadOnly = true;
            dataGridView9.Columns[18].ReadOnly = true;
            dataGridView9.Columns[21].ReadOnly = true;
            dataGridView9.Columns[22].ReadOnly = true;

            dataGridView9.Columns[1].Width = 40;
            dataGridView9.Columns[2].Width = 85;
            dataGridView9.Columns[3].Width = 230;
            dataGridView9.Columns[4].Width = 70;
            dataGridView9.Columns[5].Width = 65;
            dataGridView9.Columns[6].Width = 65;
            dataGridView9.Columns[7].Width = 65;
            dataGridView9.Columns[8].Width = 50;
            dataGridView9.Columns[9].Width = 70;
            dataGridView9.Columns[10].Width = 70;
            dataGridView9.Columns[11].Width = 70;
            dataGridView9.Columns[12].Width = 70;
            dataGridView9.Columns[13].Width = 60;
            dataGridView9.Columns[14].Width = 60;
            dataGridView9.Columns[15].Width = 70;
            dataGridView9.Columns[16].Width = 70;
            dataGridView9.Columns[17].Width = 70;
            dataGridView9.Columns[18].Width = 70;
            dataGridView9.Columns[19].Width = 85;
            dataGridView9.Columns[20].Width = 85;
            dataGridView9.Columns[21].Width = 85;
            dataGridView9.Columns[22].Width = 85;
            dataGridView9.Columns[23].Width = 85;
            dataGridView9.Columns[24].Width = 70;
            dataGridView9.Columns[25].Width = 70;
            //
            dataGridView9.EnableHeadersVisualStyles = false;
            dataGridView9.Columns[2].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[3].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[4].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[5].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[6].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[8].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[13].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[15].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[17].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[19].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[19].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[20].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[23].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[24].HeaderCell.Style.BackColor = cl_cp_header;
            dataGridView9.Columns[25].HeaderCell.Style.BackColor = cl_cp_header;
           
            foreach (DataGridViewColumn col in dataGridView9.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;

                //col.HeaderCell.Style.Font = new Font("Arial", 12F, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
            }
        }

        public void UploadCP()
        {
            string cp_id = "";
            string cp_date = "";
            string partCode = "";
            string partName = "";
            string hsCode = "";
            string quantity = "";
            string unitPrice = "";
            string amountPrice = "";
            string percentage = "";
            string weight = "";
            string priceRw = "";
            string priceAir = "";
            string poTime = "";
            string deliveryAIR = "";
            string deliveryRW = "";
            //
            string order_id = "";
            string order_date = "";

            DateTime dt = DateTime.Now;
            string s1 = dt.Day.ToString();
            string s2 = dt.Month.ToString();
            string s3 = dt.Year.ToString();
            s3 = s3.Substring(2, 2);
            string s4 = dt.Hour.ToString();
            string s5 = dt.Minute.ToString();
            string s6 = dt.Second.ToString();
            if (s1.Length == 1) { s1 = "0" + s1; }
            if (s2.Length == 1) { s2 = "0" + s2; }
            if (s4.Length == 1) { s4 = "0" + s4; }
            if (s5.Length == 1) { s5 = "0" + s5; }
            if (s6.Length == 1) { s6 = "0" + s6; }
            cp_id = "1" + s1 + s2 + s3 + s4 + s5 + s6;
            textBox15.Text = cp_id;
            //
            string year = Convert.ToString(dateTimePicker10.Value.Year);
            string month = Convert.ToString(dateTimePicker10.Value.Month);
            string day = Convert.ToString(dateTimePicker10.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            cp_date = day + "." + month + "." + year;
            //
            year = Convert.ToString(dateTimePicker15.Value.Year);
            month = Convert.ToString(dateTimePicker15.Value.Month);
            day = Convert.ToString(dateTimePicker15.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            order_date = day + "." + month + "." + year;
            order_id = textBox30.Text;

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "INSERT INTO com_proposal(cp_id, cp_date, delivery_type, rw, air, order_id, order_date, customer, project) VALUES(@cpid, @cpdate, @delivery, @rw, @air, @order, @order_dt, @customer, @project)";
            command1.Parameters.AddWithValue("@cpid", cp_id);
            command1.Parameters.AddWithValue("@cpdate", cp_date);
            command1.Parameters.AddWithValue("@delivery", comboBox3.Text);
            command1.Parameters.AddWithValue("@rw", comboBox16.Text);
            command1.Parameters.AddWithValue("@air", comboBox17.Text);
            command1.Parameters.AddWithValue("@order", order_id);
            command1.Parameters.AddWithValue("@order_dt", order_date);
            command1.Parameters.AddWithValue("@customer", comboBox14.Text);
            command1.Parameters.AddWithValue("@project", textBox35.Text);
            command1.ExecuteNonQuery();
            //
            int i = 0;
            for (i = 0; i < dataGridView9.Rows.Count-1; i++)
            {
                try { partCode = dataGridView9.Rows[i].Cells[2].Value.ToString(); } catch { partCode = ""; }
                try { partName = dataGridView9.Rows[i].Cells[3].Value.ToString(); } catch { partName = ""; }
                try { hsCode = dataGridView9.Rows[i].Cells[4].Value.ToString(); } catch { hsCode = ""; }
                try { quantity = dataGridView9.Rows[i].Cells[5].Value.ToString(); } catch { quantity = ""; }
                try { unitPrice = dataGridView9.Rows[i].Cells[6].Value.ToString(); } catch { unitPrice = ""; }
                try { amountPrice = dataGridView9.Rows[i].Cells[7].Value.ToString(); } catch { amountPrice = ""; }
                try { percentage = dataGridView9.Rows[i].Cells[8].Value.ToString(); } catch { percentage = ""; }
                try { weight = dataGridView9.Rows[i].Cells[13].Value.ToString(); } catch { weight = ""; }
                try { priceRw = dataGridView9.Rows[i].Cells[15].Value.ToString(); } catch { priceRw = ""; }
                try { priceAir = dataGridView9.Rows[i].Cells[17].Value.ToString(); } catch { priceAir = ""; }
                try { poTime = dataGridView9.Rows[i].Cells[23].Value.ToString(); } catch { poTime = ""; }
                try { deliveryAIR = dataGridView9.Rows[i].Cells[24].Value.ToString(); } catch { deliveryAIR = ""; }
                try { deliveryRW = dataGridView9.Rows[i].Cells[25].Value.ToString(); } catch { deliveryRW = ""; }

                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "INSERT INTO items_in_cp(part_code, part_name, hs_code, quantity, unit_price, amount_price, cp_id, percentage, weight_of_item, price_rw, price_air, production_time, delivery_air, delivery_rw) VALUES(@partcode, @partname, @hscode, @quantity, @unitprice, @amountprice, @cpid, @percentage, @weight, @pricerw, @priceair, @potime, @dair, @drw)";
                command.Parameters.AddWithValue("@partcode", partCode);
                command.Parameters.AddWithValue("@partname", partName);
                command.Parameters.AddWithValue("@hscode", hsCode);
                command.Parameters.AddWithValue("@quantity", quantity);
                command.Parameters.AddWithValue("@unitprice", unitPrice);
                command.Parameters.AddWithValue("@amountprice", amountPrice);
                command.Parameters.AddWithValue("@cpid", cp_id);
                command.Parameters.AddWithValue("@percentage", percentage);
                command.Parameters.AddWithValue("@weight", weight);
                command.Parameters.AddWithValue("@pricerw", priceRw);
                command.Parameters.AddWithValue("@priceair", priceAir);
                command.Parameters.AddWithValue("@potime", poTime);
                command.Parameters.AddWithValue("@dair", deliveryAIR);
                command.Parameters.AddWithValue("@drw", deliveryRW);
                command.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("Комерческое предложение сохраненно.", "Done",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
            button16.Enabled = true;
            button19.Enabled = false;
            button18.Enabled = false;
            button48.Enabled = false;
        }

        public void RefreshCPGrid(string text, int type)
        {
            if (text.Length > 0)
            {
                // Refresh %
                if (type == 0)
                {
                    for (int k = 0; k < dataGridView9.Rows.Count-1; k++)
                    {
                        if (Convert.ToBoolean(dataGridView9.Rows[k].Cells[0].Value))
                        {
                            dataGridView9.Rows[k].Cells[8].Value = Math.Round(Convert.ToDouble(textBox18.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        }
                    }
                }
                // Refresh RW
                if (type == 1)
                {
                    for (int k = 0; k < dataGridView9.Rows.Count-1; k++)
                    {
                        if (Convert.ToBoolean(dataGridView9.Rows[k].Cells[0].Value))
                        {
                            dataGridView9.Rows[k].Cells[15].Value = Math.Round(Convert.ToDouble(textBox19.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        }
                    }
                }
                // Refresh AIR
                if (type == 2)
                {
                    for (int k = 0; k < dataGridView9.Rows.Count-1; k++)
                    {
                        if (Convert.ToBoolean(dataGridView9.Rows[k].Cells[0].Value))
                        {
                            dataGridView9.Rows[k].Cells[17].Value = Math.Round(Convert.ToDouble(textBox20.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                            dataGridView9.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        }
                    }
                }
                int indexOfTotalRow = dataGridView9.RowCount - 1;
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = 0;

                dataGridView9.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

                double counter = 0;
                for (int k = 0; k < indexOfTotalRow; k++)
                {
                    dataGridView9.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                    dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    counter++;
                }
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            }
        }

        public void CPGridCellChange(int row, int column)
        {
            if (column == 8)
            {
                dataGridView9.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);

            }
            if (column == 13)
            {
                dataGridView9.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 15)
            {
                dataGridView9.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 17)
            {
                dataGridView9.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 19)
            {
                dataGridView9.Rows[row].Cells[8].Value = Math.Round(((Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) - (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) - Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) / Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) * 100, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 20)
            {
                dataGridView9.Rows[row].Cells[8].Value = Math.Round(((Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) - (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) - Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) / Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)) * 100, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 5)
            {
                dataGridView9.Rows[row].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            if (column == 6)
            {
                dataGridView9.Rows[row].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
            }
            //otmetka
            if (column == 2)
            {
                string item_code = dataGridView9.Rows[row].Cells[2].Value.ToString();

                string item_name = "";
                string unit_price = "0";
                string unit_weight = "0";
                string qty = "1";
                bool found = false;

                if(found==false)
                {
                    SqlConnection connection = new SqlConnection(conString);
                    connection.Open();
                    SqlCommand command = new SqlCommand();
                    using (connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (command = new SqlCommand("SELECT name, price, exact_weight FROM items_first_type WHERE item_code = '" + item_code + "'", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    item_name = reader.GetValue(0).ToString();
                                    unit_price = reader.GetValue(1).ToString();
                                    unit_weight = reader.GetValue(2).ToString();
                                    found = true;
                                    if (unit_price.Length <= 0)
                                    {
                                        unit_price = "0";
                                    }
                                    if (unit_weight.Length <= 0)
                                    {
                                        unit_weight = "0";
                                    }
                                }
                            }
                        }
                        connection.Close();
                    }
                }
                if(found==false)
                {
                    SqlConnection connection = new SqlConnection(conString);
                    connection.Open();
                    SqlCommand command = new SqlCommand();
                    using (connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (command = new SqlCommand("SELECT name, price, exact_weight FROM items_second_type WHERE item_code = '" + item_code + "'", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    item_name = reader.GetValue(0).ToString();
                                    unit_price = reader.GetValue(1).ToString();
                                    unit_weight = reader.GetValue(2).ToString();
                                    found = true;
                                    if(unit_price.Length<=0)
                                    {
                                        unit_price = "0";
                                    }
                                    if(unit_weight.Length<=0)
                                    {
                                        unit_weight = "0";
                                    }
                                }
                            }
                        }
                        connection.Close();
                    }
                }

                if(found==true)
                {
                    dataGridView9.Rows[row].Cells[3].Value = item_name;
                    dataGridView9.Rows[row].Cells[4].Value = "";
                    dataGridView9.Rows[row].Cells[5].Value = qty;
                    dataGridView9.Rows[row].Cells[6].Value = unit_price;
                    dataGridView9.Rows[row].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[9].Value = Math.Round((Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100)), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[10].Value = Math.Round((Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven); ;
                    dataGridView9.Rows[row].Cells[11].Value = (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven));
                    dataGridView9.Rows[row].Cells[12].Value = Math.Round((Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven); 
                    dataGridView9.Rows[row].Cells[13].Value = unit_weight;
                    dataGridView9.Rows[row].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[15].Value = 0; 
                    dataGridView9.Rows[row].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[17].Value = 7; 
                    dataGridView9.Rows[row].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[row].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[row].Cells[23].Value = "2-3 week";
                    dataGridView9.Rows[row].Cells[24].Value = "2-3 days";
                    dataGridView9.Rows[row].Cells[25].Value = "3-4 week";
                }
                else
                {
                    MessageBox.Show("cannot find item: "+dataGridView9.Rows[row].Cells[2].Value.ToString(), "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            int indexOfTotalRow = dataGridView9.RowCount - 1;
            dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = 0;

            dataGridView9.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView9.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        public void LoadCPGrid(string doc)
        {
            int indexOfTotalRow = 0;
            double defaultPercentage = 71;
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (doc == "") {
                // Open from dialog
                //
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo file = new FileInfo(openFileDialog.FileName);
                    ExcelPackage excel = new ExcelPackage(file);
                    var worksheet = excel.Workbook.Worksheets[1];

                    dataGridView9.RowCount = worksheet.Dimension.Rows;
                    indexOfTotalRow = worksheet.Dimension.Rows - 1;
                    for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                    {
                        dataGridView9.Rows[k].Cells[1].Value = k + 1;
                        var ws5 = worksheet.Cells[k + 2, 5].Value;
                        var ws6 = worksheet.Cells[k + 2, 6].Value;
                        if (ws5 == null) { ws5 = 0; }
                        if (ws6 == null) { ws6 = 0; }
                        dataGridView9.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 2].Value;
                        dataGridView9.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 3].Value;
                        dataGridView9.Rows[k].Cells[4].Value = worksheet.Cells[k + 2, 4].Value;
                        dataGridView9.Rows[k].Cells[5].Value = ws5;
                        dataGridView9.Rows[k].Cells[6].Value = ws6;
                        dataGridView9.Rows[k].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[8].Value = defaultPercentage;
                        dataGridView9.Rows[k].Cells[9].Value = Math.Round((Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100)), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[10].Value = Math.Round((Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven); ;
                        dataGridView9.Rows[k].Cells[11].Value = (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven));
                        dataGridView9.Rows[k].Cells[12].Value = Math.Round((Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven); ;
                        var ws8 = worksheet.Cells[k + 2, 8].Value;
                        if (ws8 == null) { ws8 = 0; }
                        dataGridView9.Rows[k].Cells[13].Value = ws8;
                        dataGridView9.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[15].Value = 0; // textBox19.Text
                        dataGridView9.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[17].Value = 7; // textBox20.Text
                        dataGridView9.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[k].Cells[23].Value = "2-3 week";
                        dataGridView9.Rows[k].Cells[24].Value = "2-3 days";
                        dataGridView9.Rows[k].Cells[25].Value = "3-4 week";
                        //
                        if (k % 2 == 0)
                        {
                            dataGridView9.Rows[k].Cells[0].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[1].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[2].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[3].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[4].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[5].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[6].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[7].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                            dataGridView9.Rows[k].Cells[9].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[10].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[11].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[12].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[13].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[14].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[15].Style.BackColor = Color.SandyBrown; //
                            dataGridView9.Rows[k].Cells[16].Style.BackColor = Color.SandyBrown; //
                            dataGridView9.Rows[k].Cells[17].Style.BackColor = Color.SkyBlue;
                            dataGridView9.Rows[k].Cells[18].Style.BackColor = Color.SkyBlue;
                            dataGridView9.Rows[k].Cells[19].Style.BackColor = Color.SandyBrown;
                            dataGridView9.Rows[k].Cells[20].Style.BackColor = Color.SkyBlue;
                            dataGridView9.Rows[k].Cells[21].Style.BackColor = Color.SandyBrown; //
                            dataGridView9.Rows[k].Cells[22].Style.BackColor = Color.SkyBlue; //
                            dataGridView9.Rows[k].Cells[23].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[24].Style.BackColor = cl_even;
                            dataGridView9.Rows[k].Cells[25].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView9.Rows[k].Cells[0].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[1].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[2].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[3].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[4].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[5].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[6].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[7].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                            dataGridView9.Rows[k].Cells[9].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[10].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[11].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[12].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[13].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[14].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[15].Style.BackColor = Color.SandyBrown; //
                            dataGridView9.Rows[k].Cells[16].Style.BackColor = Color.SandyBrown; //
                            dataGridView9.Rows[k].Cells[17].Style.BackColor = Color.LightSkyBlue;
                            dataGridView9.Rows[k].Cells[18].Style.BackColor = Color.LightSkyBlue;
                            dataGridView9.Rows[k].Cells[19].Style.BackColor = Color.SandyBrown;
                            dataGridView9.Rows[k].Cells[20].Style.BackColor = Color.LightSkyBlue;
                            dataGridView9.Rows[k].Cells[21].Style.BackColor = Color.SandyBrown; //
                            dataGridView9.Rows[k].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                            dataGridView9.Rows[k].Cells[23].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[24].Style.BackColor = cl_odd;
                            dataGridView9.Rows[k].Cells[25].Style.BackColor = cl_odd;
                            // SeaGreen MediumSeaGreen DarkSeaGreen
                            // AliceBlue CornflowerBlue AliceBlue
                        }
                    }
                    dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = 0;
                    dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = 0;

                    dataGridView9.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
                    dataGridView9.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
                    dataGridView9.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
                    dataGridView9.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
                    dataGridView9.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
                    dataGridView9.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
                    dataGridView9.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

                    double counter = 0;
                    for (int k = 0; k < indexOfTotalRow; k++)
                    {
                        dataGridView9.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                        dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                        counter++;
                    }
                    dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                }
            } else {
                // open from template
                //
                FileInfo file = new FileInfo(doc);
                ExcelPackage excel = new ExcelPackage(file);
                var worksheet = excel.Workbook.Worksheets[1];

                dataGridView9.RowCount = worksheet.Dimension.Rows;
                indexOfTotalRow = worksheet.Dimension.Rows - 1;
                //
                //if (textBox19.Text == "") { textBox19.Text = "0"; }
                //if (textBox20.Text == "") { textBox20.Text = "0"; }
                //
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    //dataGridView9.Rows[k].Cells[1].Value = worksheet.Cells[k + 2, 1].Value;
                    dataGridView9.Rows[k].Cells[1].Value = k + 1;

                    var ws5 = worksheet.Cells[k + 2, 5].Value;
                    var ws6 = worksheet.Cells[k + 2, 6].Value;
                    if (ws5 == null) { ws5 = 0; }
                    if (ws6 == null) { ws6 = 0; }
                    dataGridView9.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView9.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView9.Rows[k].Cells[4].Value = worksheet.Cells[k + 2, 4].Value;
                    dataGridView9.Rows[k].Cells[5].Value = ws5;
                    dataGridView9.Rows[k].Cells[6].Value = ws6;

                    //dataGridView9.Rows[k].Cells[7].Value = worksheet.Cells[k + 2, 7].Value;
                    dataGridView9.Rows[k].Cells[7].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[8].Value = defaultPercentage;
                    dataGridView9.Rows[k].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / 100), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[10].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[12].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    var ws8 = worksheet.Cells[k + 2, 8].Value;
                    if (ws8 == null) { ws8 = 0; }
                    dataGridView9.Rows[k].Cells[13].Value = ws8;
                    dataGridView9.Rows[k].Cells[14].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[15].Value = 0; // textBox19.Text
                    dataGridView9.Rows[k].Cells[16].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[17].Value = 7; // textBox20.Text
                    dataGridView9.Rows[k].Cells[18].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + (Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven)), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[21].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[22].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) * Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[k].Cells[23].Value = "2-3 week";
                    dataGridView9.Rows[k].Cells[24].Value = "2-3 days";
                    dataGridView9.Rows[k].Cells[25].Value = "3-4 week";
                    //
                    if (k % 2 == 0)
                    {
                        dataGridView9.Rows[k].Cells[0].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[1].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[2].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[3].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[4].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[5].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[6].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[7].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                        dataGridView9.Rows[k].Cells[9].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[10].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[11].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[12].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[13].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[14].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[15].Style.BackColor = Color.SandyBrown; //
                        dataGridView9.Rows[k].Cells[16].Style.BackColor = Color.SandyBrown; //
                        dataGridView9.Rows[k].Cells[17].Style.BackColor = Color.SkyBlue;
                        dataGridView9.Rows[k].Cells[18].Style.BackColor = Color.SkyBlue;
                        dataGridView9.Rows[k].Cells[19].Style.BackColor = Color.SandyBrown;
                        dataGridView9.Rows[k].Cells[20].Style.BackColor = Color.SkyBlue;
                        dataGridView9.Rows[k].Cells[21].Style.BackColor = Color.SandyBrown; //
                        dataGridView9.Rows[k].Cells[22].Style.BackColor = Color.SkyBlue; //
                        dataGridView9.Rows[k].Cells[23].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[24].Style.BackColor = cl_even;
                        dataGridView9.Rows[k].Cells[25].Style.BackColor = cl_even;
                    }
                    else
                    {
                        dataGridView9.Rows[k].Cells[0].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[1].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[2].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[3].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[4].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[5].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[6].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[7].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                        dataGridView9.Rows[k].Cells[9].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[10].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[11].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[12].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[13].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[14].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[15].Style.BackColor = Color.SandyBrown; //
                        dataGridView9.Rows[k].Cells[16].Style.BackColor = Color.SandyBrown; //
                        dataGridView9.Rows[k].Cells[17].Style.BackColor = Color.LightSkyBlue;
                        dataGridView9.Rows[k].Cells[18].Style.BackColor = Color.LightSkyBlue;
                        dataGridView9.Rows[k].Cells[19].Style.BackColor = Color.SandyBrown;
                        dataGridView9.Rows[k].Cells[20].Style.BackColor = Color.LightSkyBlue;
                        dataGridView9.Rows[k].Cells[21].Style.BackColor = Color.SandyBrown; //
                        dataGridView9.Rows[k].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                        dataGridView9.Rows[k].Cells[23].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[24].Style.BackColor = cl_odd;
                        dataGridView9.Rows[k].Cells[25].Style.BackColor = cl_odd;
                        // SeaGreen MediumSeaGreen DarkSeaGreen
                        // AliceBlue CornflowerBlue AliceBlue
                    }
                }
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = 0;
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = 0;

                dataGridView9.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
                dataGridView9.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
                dataGridView9.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

                double counter = 0;
                for (int k = 0; k < indexOfTotalRow; k++)
                {
                    dataGridView9.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                    dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    counter++;
                }
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            }
            dataGridView9.ClearSelection();
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
            command.CommandText = "UPDATE Boxes SET black_id='" + grey_id + "' WHERE black_id='" + black_id + "'";
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

        public string MakeGreyID()
        {
            // Генерация Grey ID
            string grayId;
            Random rnd = new Random();
            int valueFirst = rnd.Next(999, 9999);
            int valueSecond = rnd.Next(999, 9999);
            int valueThird = rnd.Next(999, 9999);
            grayId = valueFirst.ToString() + valueSecond.ToString() + valueThird.ToString();
            return grayId;
        }

        public void GenerateGreyID(int qty, string type)
        {
            // Генерация Grey ID
            Document document = new Document();
            //
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "DOCX (*.docx)|*.docx";
            saveFileDialog.FilterIndex = 3;
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            int i = 0;
            for (i = 0; i < qty; i++)
            {
                string grayIdToPrint = "";
                bool checkGreyId = false;
                while (checkGreyId == false)
                {
                    grayIdToPrint = MakeGreyID();
                    SqlCommand check_Grey_Id = new SqlCommand("SELECT COUNT(*) FROM [Boxes] WHERE ([grey_id] = @greyid)", connection);
                    check_Grey_Id.Parameters.AddWithValue("@greyid", grayIdToPrint);
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
                var resultImage = new Bitmap(grayIdToPrint.ToString().Length * 40, 150);
                var resultText = new Bitmap(580, 95);
                using (Graphics graphics = Graphics.FromImage(resultImage))
                {
                    Font oFont = new System.Drawing.Font("IDAutomationHC39M", 25);
                    PointF point = new PointF(2f, 2f);
                    SolidBrush black = new SolidBrush(Color.Black);
                    SolidBrush white = new SolidBrush(cl_odd);
                    graphics.FillRectangle(white, 0, 0, resultImage.Width, resultImage.Height);
                    graphics.DrawString("*" + grayIdToPrint + "*", oFont, black, point);
                }
                using (Graphics graphicsText = Graphics.FromImage(resultText))
                {
                    Font tFont = new System.Drawing.Font("Microsoft Sans Serif", 60);
                    PointF pointText = new PointF(2f, 2f);
                    SolidBrush black = new SolidBrush(Color.Black);
                    SolidBrush white = new SolidBrush(cl_odd);
                    graphicsText.FillRectangle(white, 0, 0, resultText.Width, resultText.Height);
                    graphicsText.DrawString(grayIdToPrint, tFont, black, pointText);
                }
                Section section = document.AddSection();
                section.TextDirection = Spire.Doc.Documents.TextDirection.LeftToRightRotated;
                Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

                DocPicture picture = document.Sections[i].Paragraphs[0].AppendPicture(resultImage);
                DocPicture pictureText = document.Sections[i].Paragraphs[0].AppendPicture(resultText);

                picture.HorizontalOrigin = HorizontalOrigin.Page;
                //picture.HorizontalAlignment = ShapeHorizontalAlignment.Right;
                picture.VerticalOrigin = VerticalOrigin.Page;
                picture.VerticalAlignment = ShapeVerticalAlignment.Center;
                picture.TextWrappingStyle = TextWrappingStyle.Through;

                pictureText.HorizontalOrigin = HorizontalOrigin.Page;
                //pictureText.HorizontalAlignment = ShapeHorizontalAlignment.Left;
                pictureText.VerticalOrigin = VerticalOrigin.Page;
                pictureText.VerticalAlignment = ShapeVerticalAlignment.Center;
                pictureText.TextWrappingStyle = TextWrappingStyle.Through;

                pictureText.HorizontalPosition = -15.0F;
                picture.HorizontalPosition = 180.0F;

                pictureText.Rotation = 90;
                picture.Rotation = 90;
                section.TextDirection = Spire.Doc.Documents.TextDirection.RightToLeftRotated;
            }
            //
            if (type == "print") {
                document.SaveToFile(app_dir_temp + "grey_id.docx");
                //
                PrintDialog printDlg = new PrintDialog();
                printDlg.AllowPrintToFile = true;
                //printDlg.AllowCurrentPage = true;
                //printDlg.AllowSelection = true;
                //printDlg.AllowSomePages = true;
                printDlg.UseEXDialog = true;
                //
                Document Doc = new Document();
                Doc.LoadFromFile(app_dir_temp + "grey_id.docx");
                Doc.PrintDialog = printDlg;
                //
                System.Drawing.Printing.PrintDocument printDoc = Doc.PrintDocument;
                //printDoc.Print();
                //
                if (printDlg.ShowDialog() == DialogResult.OK) {
                    printDoc.Print();
                }
                File.Delete(app_dir_temp + "grey_id.docx");
                //
            } else {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    document.SaveToFile(saveFileDialog.FileName);
                }
            }
            //
        }

        public class Item
        {
            public string boxNumber { get; set; }
            public string black_id { get; set; }
            public string partName { get; set; }
            public string partCode { get; set; }
            public string amount { get; set; }
            public string netWeight { get; set; }
            public string grossWeight { get; set; }
        }

        public void GeneratePL(string idOrder, string type)
        {
            List<Item> item = new List<Item>();
            Document document = new Document();
            //int i = 0;
            //string idOrder = textBox2.Text.ToString();
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT box_numb, part_name, part_code, amount, net, gross, black_id FROM items_in_boxes WHERE id_order=@orderId ORDER BY box_numb";
            comm.Parameters.AddWithValue("@orderId", idOrder);
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                item.Add(new Item()
                {
                    boxNumber = Convert.ToString(reader["box_numb"]),
                    black_id = Convert.ToString(reader["black_id"]),
                    partName = Convert.ToString(reader["part_name"]),
                    partCode = Convert.ToString(reader["part_code"]),
                    amount = Convert.ToString(reader["amount"]),
                    netWeight = Convert.ToString(reader["net"]),
                    grossWeight = Convert.ToString(reader["gross"])
                });
            }
            MakePLDoc(item, idOrder, type);
        }

        public void MakePLDoc(List<Item> items, string idOrder, string type)
        {
            int pageCount = 1;
            int totalBoxes = 0;
            float totalNet = 0;
            float totalGross = 0;
            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");

            worksheet.Row(1).Height = 50;
            worksheet.Column(1).Width = 1;
            worksheet.Column(2).Width = 2;
            worksheet.Column(3).Width = 2;
            worksheet.Column(4).Width = 2;
            worksheet.Column(5).Width = 3;
            worksheet.Column(6).Width = 2;
            worksheet.Column(7).Width = 3;
            worksheet.Column(8).Width = 2;
            worksheet.Column(9).Width = 3;
            worksheet.Column(10).Width = 3;
            worksheet.Column(11).Width = 3;
            worksheet.Column(12).Width = 3;
            worksheet.Column(13).Width = 2;
            worksheet.Column(14).Width = 2;
            worksheet.Column(15).Width = 2;
            worksheet.Column(16).Width = 2;
            worksheet.Column(17).Width = 2;
            worksheet.Column(18).Width = 3;
            worksheet.Column(19).Width = 2;
            worksheet.Column(20).Width = 2;
            worksheet.Column(21).Width = 2;
            worksheet.Column(22).Width = 2;
            worksheet.Column(23).Width = 2;
            worksheet.Column(24).Width = 3;
            worksheet.Column(25).Width = 2;
            worksheet.Column(26).Width = 2;
            worksheet.Column(27).Width = 3;
            worksheet.Column(27).Width = 4;
            worksheet.Column(28).Width = 3;
            worksheet.Column(29).Width = 3;

            worksheet.Row(2).Height = 25;
            worksheet.Cells[2, 10].Value = "PACKING LIST";

            using (ExcelRange rng = worksheet.Cells[4, 2, 4, 31])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[11, 2, 11, 31])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[17, 20, 17, 31])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[18, 2, 18, 19])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[26, 2, 26, 19])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[29, 2, 29, 19])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[32, 2, 32, 31])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[4, 19, 31, 19])
            {
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[26, 10, 31, 10])
            {
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }




            worksheet.Cells[4, 2].Value = "Shipper/Exporter";
            worksheet.Cells[5, 3].Value = "LSIS CO., LTD";
            worksheet.Cells[6, 3].Value = "LS Tower 1026-6, Hogye-dong, Dongan-gu,";
            worksheet.Cells[7, 3].Value = "Anyang-si, Gyeonggi-do, 431-848, Korea.";
            worksheet.Cells[8, 3].Value = "SANG-KU KANG";
            worksheet.Cells[9, 3].Value = "TEL 82 2 2034 4429";
            worksheet.Cells[4, 20].Value = "No. & date of invoice";
            worksheet.Cells[5, 21].Value = "LEO-1708012";
            worksheet.Cells[5, 28].Value = "DATED";
            worksheet.Cells[5, 30].Value = "AUG.04,2017";
            worksheet.Cells[7, 20].Value = "No. & date of L/C";
            worksheet.Cells[8, 21].Value = "NO of L/C";
            worksheet.Cells[8, 28].Value = "DATED";
            worksheet.Cells[8, 30].Value = "AUG.04,2017";
            worksheet.Cells[11, 2].Value = "For account & risk of Messers.";
            worksheet.Cells[12, 3].Value = "미티오 글로벌.LTD";
            worksheet.Cells[13, 3].Value = "경기도 안양시 동안구 관양동 810";
            worksheet.Cells[14, 3].Value = "금강펜테리움 IT 타워";
            worksheet.Cells[15, 3].Value = "남호용";
            worksheet.Cells[16, 3].Value = "TEL: 031-337-6248";
            worksheet.Cells[17, 3].Value = "H.P: 010-3356-1874";
            worksheet.Cells[11, 20].Value = "L/C issuing bank";
            worksheet.Cells[17, 20].Value = "Remarks";
            worksheet.Cells[18, 2].Value = "Notify Party";
            worksheet.Cells[19, 3].Value = "Same as above";
            worksheet.Cells[26, 2].Value = "Port of Loading";
            worksheet.Cells[27, 3].Value = "BUSAN,  KOREA";
            worksheet.Cells[26, 11].Value = "Final Destination";
            worksheet.Cells[27, 12].Value = "TASHKENT,";
            worksheet.Cells[28, 12].Value = "UZBEKISTAN";
            worksheet.Cells[29, 2].Value = "Carrier";
            worksheet.Cells[29, 11].Value = "Sailing on or about";

            using (ExcelRange rng = worksheet.Cells[46, 2, 46, 31])
            {
                rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            worksheet.Cells[47, 19].Value = pageCount.ToString();
            pageCount++;

            worksheet.Row(48).Height = 25;

            worksheet.Cells[48, 2].Value = "Box NO";
            worksheet.Cells[48, 2, 48, 4].Merge = true;

            worksheet.Cells[48, 5].Value = "Description of Goods";
            worksheet.Cells[48, 5, 48, 12].Merge = true;

            worksheet.Cells[48, 13].Value = "Item Code";
            worksheet.Cells[48, 13, 48, 17].Merge = true;

            worksheet.Cells[48, 18].Value = "(EA)";
            worksheet.Cells[48, 18, 48, 18].Merge = true;

            worksheet.Cells[48, 19].Value = "QTY";
            worksheet.Cells[48, 19, 48, 20].Merge = true;

            worksheet.Cells[48, 21].Value = "Net";
            worksheet.Cells[48, 21, 48, 23].Merge = true;

            worksheet.Cells[48, 24].Value = "Gross";
            worksheet.Cells[48, 24, 48, 26].Merge = true;

            worksheet.Cells[48, 27].Value = "Measurement";
            worksheet.Cells[48, 27, 48, 28].Merge = true;

            worksheet.Cells[48, 29].Value = "Barcode";
            worksheet.Cells[48, 29, 48, 31].Merge = true;

            using (ExcelRange rng = worksheet.Cells[48, 2, 48, 31])
            {
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            }

            int subAmount = 0;
            float subNetWeight = 0;
            float subGrossWeight = 0;
            int rowCount = 49;
            int height = 1;
            string boxNumberKeep = "0";
            for (int i = 0; i < items.Count; i++)
            {
                if (height == 45)
                {
                    rowCount = rowCount + 4;
                    worksheet.Cells[rowCount, 19].Value = pageCount.ToString();
                    pageCount++;
                    rowCount++;

                    height = 1;
                    worksheet.Row(rowCount).Height = 25;

                    worksheet.Cells[rowCount, 2].Value = "Box No";
                    worksheet.Cells[rowCount, 2, rowCount, 4].Merge = true;

                    worksheet.Cells[rowCount, 5].Value = "Description of Goods";
                    worksheet.Cells[rowCount, 5, rowCount, 12].Merge = true;

                    worksheet.Cells[rowCount, 13].Value = "Item Code";
                    worksheet.Cells[rowCount, 13, rowCount, 17].Merge = true;

                    worksheet.Cells[rowCount, 18].Value = "(EA)";
                    worksheet.Cells[rowCount, 18, rowCount, 18].Merge = true;

                    worksheet.Cells[rowCount, 19].Value = "QTY";
                    worksheet.Cells[rowCount, 19, rowCount, 20].Merge = true;

                    worksheet.Cells[rowCount, 21].Value = "Net";
                    worksheet.Cells[rowCount, 21, rowCount, 23].Merge = true;

                    worksheet.Cells[rowCount, 24].Value = "Gross";
                    worksheet.Cells[rowCount, 24, rowCount, 26].Merge = true;

                    worksheet.Cells[rowCount, 27].Value = "Measurement";
                    worksheet.Cells[rowCount, 27, rowCount, 28].Merge = true;

                    worksheet.Cells[rowCount, 29].Value = "Barcode";
                    worksheet.Cells[rowCount, 29, rowCount, 31].Merge = true;

                    using (ExcelRange rng = worksheet.Cells[rowCount, 2, rowCount, 31])
                    {
                        rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    }

                    rowCount++;
                }
                if (boxNumberKeep != items[i].boxNumber)
                {
                    if (subAmount > 0)
                    {
                        worksheet.Cells[rowCount, 5, rowCount, 13].Merge = true;
                        worksheet.Cells[rowCount, 5].Value = "Sub Total: ";
                        worksheet.Cells[rowCount, 19, rowCount, 20].Merge = true;
                        worksheet.Cells[rowCount, 19].Value = subAmount.ToString();
                        worksheet.Cells[rowCount, 21, rowCount, 23].Merge = true;
                        worksheet.Cells[rowCount, 21].Value = subNetWeight.ToString();
                        worksheet.Cells[rowCount, 24, rowCount, 26].Merge = true;
                        worksheet.Cells[rowCount, 24].Value = subGrossWeight.ToString();
                        worksheet.Cells[rowCount, 27, rowCount, 28].Merge = true;
                        using (ExcelRange rng = worksheet.Cells[rowCount, 31, rowCount, 31])
                        {
                            rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }
                        using (ExcelRange rng = worksheet.Cells[rowCount, 2, rowCount, 4])
                        {
                            rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }
                        using (ExcelRange rng = worksheet.Cells[rowCount, 27, rowCount, 28])
                        {
                            rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        using (ExcelRange rng = worksheet.Cells[rowCount, 5, rowCount, 31])
                        {
                            rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        totalNet = totalNet + subNetWeight;
                        totalGross = totalGross + subGrossWeight;
                        rowCount++;
                        height++;
                        subAmount = 0;
                        subGrossWeight = 0;
                        subNetWeight = 0;

                        if (height == 45)
                        {
                            rowCount = rowCount + 4;
                            worksheet.Cells[rowCount, 19].Value = pageCount.ToString();
                            pageCount++;
                            rowCount++;

                            height = 1;
                            worksheet.Row(rowCount).Height = 25;

                            worksheet.Cells[rowCount, 2].Value = "Box No";
                            worksheet.Cells[rowCount, 2, rowCount, 4].Merge = true;

                            worksheet.Cells[rowCount, 5].Value = "Description of Goods";
                            worksheet.Cells[rowCount, 5, rowCount, 12].Merge = true;

                            worksheet.Cells[rowCount, 13].Value = "Item Code";
                            worksheet.Cells[rowCount, 13, rowCount, 17].Merge = true;

                            worksheet.Cells[rowCount, 18].Value = "(EA)";
                            worksheet.Cells[rowCount, 18, rowCount, 18].Merge = true;

                            worksheet.Cells[rowCount, 19].Value = "QTY";
                            worksheet.Cells[rowCount, 19, rowCount, 20].Merge = true;

                            worksheet.Cells[rowCount, 21].Value = "Net";
                            worksheet.Cells[rowCount, 21, rowCount, 23].Merge = true;

                            worksheet.Cells[rowCount, 24].Value = "Gross";
                            worksheet.Cells[rowCount, 24, rowCount, 26].Merge = true;

                            worksheet.Cells[rowCount, 27].Value = "Measurement";
                            worksheet.Cells[rowCount, 27, rowCount, 28].Merge = true;

                            worksheet.Cells[rowCount, 29].Value = "Barcode";
                            worksheet.Cells[rowCount, 29, rowCount, 31].Merge = true;

                            using (ExcelRange rng = worksheet.Cells[rowCount, 2, rowCount, 31])
                            {
                                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                            }

                            rowCount++;
                        }
                    }

                    totalBoxes++;

                    worksheet.Row(rowCount).Height = 15;
                    worksheet.Cells[rowCount, 2].Value = items[i].boxNumber.ToString();
                    worksheet.Cells[rowCount, 2, rowCount, 4].Merge = true;

                    worksheet.Cells[rowCount, 5].Value = items[i].partName.ToString();
                    worksheet.Cells[rowCount, 5, rowCount, 12].Merge = true;

                    worksheet.Cells[rowCount, 13].Value = items[i].partCode.ToString();
                    worksheet.Cells[rowCount, 13, rowCount, 17].Merge = true;

                    worksheet.Cells[rowCount, 18].Value = "EA";
                    worksheet.Cells[rowCount, 18, rowCount, 18].Merge = true;

                    worksheet.Cells[rowCount, 19].Value = items[i].amount.ToString();
                    worksheet.Cells[rowCount, 19, rowCount, 20].Merge = true;

                    worksheet.Cells[rowCount, 21].Value = items[i].netWeight.ToString();
                    worksheet.Cells[rowCount, 21, rowCount, 23].Merge = true;

                    worksheet.Cells[rowCount, 24].Value = items[i].grossWeight.ToString();
                    worksheet.Cells[rowCount, 24, rowCount, 26].Merge = true;

                    worksheet.Cells[rowCount, 27, rowCount, 28].Merge = true;

                    worksheet.Cells[rowCount, 29].Value = "*" + items[i].black_id.ToString() + "*";
                    using (ExcelRange rng = worksheet.Cells[rowCount, 29, rowCount + 1, 31])
                    {
                        rng.Style.Font.Name = "IDAutomationHC39M";
                        rng.Style.Font.Size = 10;
                        rng.Merge = true;
                        rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }


                    using (ExcelRange rng = worksheet.Cells[rowCount, 5, rowCount, 28])
                    {
                        rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    using (ExcelRange rng = worksheet.Cells[rowCount, 29, rowCount, 31])
                    {
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    using (ExcelRange rng = worksheet.Cells[rowCount, 2, rowCount, 4])
                    {
                        rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    subAmount = subAmount + Convert.ToInt32(items[i].amount);
                    subGrossWeight = subGrossWeight + Convert.ToInt32(items[i].grossWeight);
                    subNetWeight = subNetWeight + Convert.ToInt32(items[i].netWeight);
                    boxNumberKeep = items[i].boxNumber.ToString();
                    rowCount++;
                    height++;
                }
                else
                {
                    worksheet.Cells[rowCount, 2, rowCount, 4].Merge = true;
                    worksheet.Cells[rowCount, 5].Value = items[i].partName.ToString();
                    worksheet.Cells[rowCount, 5, rowCount, 12].Merge = true;

                    worksheet.Cells[rowCount, 13].Value = items[i].partCode.ToString();
                    worksheet.Cells[rowCount, 13, rowCount, 17].Merge = true;

                    worksheet.Cells[rowCount, 18].Value = "EA";
                    worksheet.Cells[rowCount, 18, rowCount, 18].Merge = true;

                    worksheet.Cells[rowCount, 19].Value = items[i].amount.ToString();
                    worksheet.Cells[rowCount, 19, rowCount, 20].Merge = true;

                    worksheet.Cells[rowCount, 21].Value = items[i].netWeight.ToString();
                    worksheet.Cells[rowCount, 21, rowCount, 23].Merge = true;

                    worksheet.Cells[rowCount, 24].Value = items[i].grossWeight.ToString();
                    worksheet.Cells[rowCount, 24, rowCount, 26].Merge = true;

                    worksheet.Cells[rowCount, 27, rowCount, 28].Merge = true;


                    using (ExcelRange rng = worksheet.Cells[rowCount, 5, rowCount, 28])
                    {
                        rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    using (ExcelRange rng = worksheet.Cells[rowCount, 31, rowCount, 31])
                    {
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    using (ExcelRange rng = worksheet.Cells[rowCount, 2, rowCount, 4])
                    {
                        rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    subAmount = subAmount + Convert.ToInt32(items[i].amount);
                    subGrossWeight = subGrossWeight + Convert.ToInt32(items[i].grossWeight);
                    subNetWeight = subNetWeight + Convert.ToInt32(items[i].netWeight);

                    rowCount++;
                    height++;
                }
                if (i == items.Count - 1)
                {
                    worksheet.Cells[rowCount, 5, rowCount, 13].Merge = true;
                    worksheet.Cells[rowCount, 5].Value = "Sub Total: ";
                    worksheet.Cells[rowCount, 19, rowCount, 20].Merge = true;
                    worksheet.Cells[rowCount, 19].Value = subAmount.ToString();
                    worksheet.Cells[rowCount, 21, rowCount, 23].Merge = true;
                    worksheet.Cells[rowCount, 21].Value = subNetWeight.ToString();
                    worksheet.Cells[rowCount, 24, rowCount, 26].Merge = true;
                    worksheet.Cells[rowCount, 24].Value = subGrossWeight.ToString();
                    worksheet.Cells[rowCount, 27, rowCount, 28].Merge = true;

                    using (ExcelRange rng = worksheet.Cells[rowCount, 31, rowCount, 31])
                    {
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    using (ExcelRange rng = worksheet.Cells[rowCount, 2, rowCount, 4])
                    {
                        rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }
                    using (ExcelRange rng = worksheet.Cells[rowCount, 27, rowCount, 28])
                    {
                        rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    using (ExcelRange rng = worksheet.Cells[rowCount, 5, rowCount, 28])
                    {
                        rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    totalNet = totalNet + subNetWeight;
                    totalGross = totalGross + subGrossWeight;
                    rowCount++;
                    height++;
                    subAmount = 0;
                    subGrossWeight = 0;
                    subNetWeight = 0;

                    worksheet.Cells[rowCount + 49 - height, 19].Value = pageCount.ToString();
                    pageCount++;
                    rowCount++;
                }
            }

            worksheet.Cells[33, 7].Value = "TOTAL: " + totalBoxes.ToString() + " BOXES,     " + "Net_Weight: " + totalNet.ToString() + "     Gross-Weight: " + totalGross.ToString();
            worksheet.Cells[33, 7, 33, 30].Merge = true;

            using (ExcelRange rng = worksheet.Cells[2, 10, 2, 24])
            {
                rng.Merge = true;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.Font.Size = 20;
                rng.Style.Font.Name = "Arial Unicode MS";
                rng.Style.WrapText = true;
            }
            using (ExcelRange rng = worksheet.Cells[1, 1, worksheet.Dimension.Rows, 28])
            {
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            }
            using (ExcelRange rng = worksheet.Cells[45, 1, worksheet.Dimension.Rows, 28])
            {
                rng.Style.WrapText = true;
                rng.Style.Font.Size = 6;
            }
            using (ExcelRange rng = worksheet.Cells[4, 2, 45, 31])
            {
                rng.Style.Font.Size = 8;
            }
          
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                excel.SaveAs(new FileInfo(saveFileDialog.FileName));
            }
        }

        //internal void GenerateGreyID()
        //{
        //    throw new NotImplementedException();
        //}

        public void GetPackingList(string param, string type)
        {
            // Функция отображает список заказаов в разделе 'История'
            //
            List<string[]> list = new List<string[]>();
            //
            if (param == "" && type != "Дата invoice" && type != "Дата packing list")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, pl_id, pl_date, invoice_id, invoice_date FROM items_in_boxes JOIN orders on items_in_boxes.id_order=orders.id_order", connection))
                    //using (SqlCommand command = new SqlCommand("SELECT DISTINCT id_order, dz_date, customer_order_id, c_date, descr, com_proposal, contract FROM orders", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "PO")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, pl_id, pl_date, invoice_id, invoice_date FROM items_in_boxes JOIN orders on items_in_boxes.id_order = orders.id_order WHERE items_in_boxes.id_order='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "Order")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, items_in_boxes.pl_id, items_in_boxes.pl_date, items_in_boxes.invoice_id, items_in_boxes.invoice_date FROM orders JOIN items_in_boxes on orders.id_order = items_in_boxes.id_order WHERE orders.customer_order_id='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "Packing list")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, items_in_boxes.pl_id, items_in_boxes.pl_date, items_in_boxes.invoice_id, items_in_boxes.invoice_date FROM orders JOIN items_in_boxes on orders.id_order = items_in_boxes.id_order WHERE items_in_boxes.pl_id = '" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //            
            if (param != "" && type == "Дата packing list")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, items_in_boxes.pl_id, items_in_boxes.pl_date, items_in_boxes.invoice_id, items_in_boxes.invoice_date FROM orders JOIN items_in_boxes on orders.id_order = items_in_boxes.id_order WHERE items_in_boxes.pl_date = '" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //Заказ от (DZ/Клиент)
            if (param != "" && type == "Invoice")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, items_in_boxes.pl_id, items_in_boxes.pl_date, items_in_boxes.invoice_id, items_in_boxes.invoice_date FROM orders JOIN items_in_boxes on orders.id_order = items_in_boxes.id_order WHERE items_in_boxes.invoice_id = '" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "Дата invoice")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT items_in_boxes.id_order, orders.customer_order_id, items_in_boxes.pl_id, items_in_boxes.pl_date, items_in_boxes.invoice_id, items_in_boxes.invoice_date FROM orders JOIN items_in_boxes on orders.id_order = items_in_boxes.id_order WHERE items_in_boxes.invoice_date = '" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[6];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            //
            dataGridView7.RowCount = 0;
            dataGridView8.RowCount = 0;
            dataGridView7.RowCount = list.Count;
            dataGridView7.RowHeadersWidth = 35;
            for (int i = 0; i < list.Count; i++)
            {
                dataGridView7.Rows[i].Cells[0].Value = list[i][0];
                dataGridView7.Rows[i].Cells[1].Value = list[i][1];
                dataGridView7.Rows[i].Cells[2].Value = list[i][2];
                dataGridView7.Rows[i].Cells[3].Value = list[i][3];
                dataGridView7.Rows[i].Cells[4].Value = list[i][4];
                dataGridView7.Rows[i].Cells[5].Value = list[i][5];
                if (i % 2 == 0)
                {
                    dataGridView7.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView7.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView7.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView7.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView7.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView7.Rows[i].Cells[5].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView7.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView7.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView7.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView7.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView7.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView7.Rows[i].Cells[5].Style.BackColor = cl_odd;
                }
            }
            //
            //Size size = new Size();
            //size.Width = dataGridView4.Columns[0].Width + dataGridView4.Columns[1].Width + dataGridView4.Columns[2].Width + dataGridView4.Columns[3].Width + dataGridView4.Columns[4].Width + dataGridView4.Columns[5].Width + dataGridView4.Columns[6].Width;
            //dataGridView4.Size = size;
        }

        public void GetOrdersList(string param, string type) {
            // Функция отображает список заказаов в разделе 'История'
            //
            List<string[]> list = new List<string[]>();
            //
            if (param == "" && type != "PO дата" && type != "Order дата") {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "PO")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE id_order='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "PO дата")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE dz_date='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "Order")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE customer_order_id='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //            
            if (param != "" && type == "Order дата")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE c_date='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //Заказ от (DZ/Клиент)
            if (param != "" && type == "Заказ от (DZ/Клиент)")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE descr='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "CP")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE com_proposal='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            if (param != "" && type == "CT")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE contract='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //            
            if (param != "" && type == "CP дата")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE comp_proposal_date='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //            
            if (param != "" && type == "CT дата")
            {
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT DISTINCT descr, id_order, dz_date, customer_order_id, c_date, com_proposal, com_proposal_date, contract, contract_date FROM orders WHERE contract_date='" + param + "'", connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string[] res_arr = new string[9];
                                res_arr[0] = reader.GetValue(0).ToString();
                                res_arr[1] = reader.GetValue(1).ToString();
                                res_arr[2] = reader.GetValue(2).ToString();
                                res_arr[3] = reader.GetValue(3).ToString();
                                res_arr[4] = reader.GetValue(4).ToString();
                                res_arr[5] = reader.GetValue(5).ToString();
                                res_arr[6] = reader.GetValue(6).ToString();
                                res_arr[7] = reader.GetValue(7).ToString();
                                res_arr[8] = reader.GetValue(8).ToString();
                                list.Add(res_arr);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            //
            richTextBox2.Clear();
            dataGridView4.RowCount = list.Count;
            dataGridView5.RowCount = 0;
            dataGridView4.RowHeadersWidth = 35;
            for (int i = 0; i < list.Count; i++) {
                dataGridView4.Rows[i].Cells[0].Value = i + 1;
                dataGridView4.Rows[i].Cells[1].Value = list[i][0];
                dataGridView4.Rows[i].Cells[2].Value = list[i][1];
                dataGridView4.Rows[i].Cells[3].Value = list[i][2];
                dataGridView4.Rows[i].Cells[4].Value = list[i][3];
                //if (list[i][3] == "client") {
                //    dataGridView4.Rows[i].Cells[4].Value = "Заказчик";
                //}
                //else {
                //dataGridView4.Rows[i].Cells[4].Value = "Drone Zone";
                //}
                dataGridView4.Rows[i].Cells[5].Value = list[i][4];
                dataGridView4.Rows[i].Cells[6].Value = list[i][5];
                dataGridView4.Rows[i].Cells[7].Value = list[i][6];
                dataGridView4.Rows[i].Cells[8].Value = list[i][7];
                dataGridView4.Rows[i].Cells[9].Value = list[i][8];

                if (i % 2 == 0) {
                    dataGridView4.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[5].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[6].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[7].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[8].Style.BackColor = cl_even;
                    dataGridView4.Rows[i].Cells[9].Style.BackColor = cl_even;
                }
                else {
                    dataGridView4.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[5].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[6].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[7].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[8].Style.BackColor = cl_odd;
                    dataGridView4.Rows[i].Cells[9].Style.BackColor = cl_odd;
                }
            }
            //
            //Size size = new Size();
            //size.Width = dataGridView4.Columns[0].Width + dataGridView4.Columns[1].Width + dataGridView4.Columns[2].Width + dataGridView4.Columns[3].Width + dataGridView4.Columns[4].Width + dataGridView4.Columns[5].Width + dataGridView4.Columns[6].Width;
            //dataGridView4.Size = size;
        }

        public void GetVersionsList(string param)
        {
            // Функция отображает список версий в соответствии с выбранным номером заказа в разделе 'История'
            //
            List<string[]> list = new List<string[]>();
            //
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT DISTINCT version, version_date from orderItems WHERE id_order='" + param + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string[] res_arr = new string[3];
                            res_arr[0] = reader.GetValue(0).ToString();
                            res_arr[1] = reader.GetValue(1).ToString();
                            list.Add(res_arr);
                        }
                    }
                }
                connection.Close();
            }
            //
            dataGridView5.RowCount = list.Count;
            for (int i = 0; i < list.Count; i++)
            {
                dataGridView5.Rows[i].Cells[0].Value = list[i][0];
                dataGridView5.Rows[i].Cells[1].Value = list[i][1];
                if (i % 2 == 0)
                {
                    dataGridView5.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView5.Rows[i].Cells[1].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView5.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView5.Rows[i].Cells[1].Style.BackColor = cl_odd;
                }
            }
        }

        public void LoadPackListItems(string id_order)
        {
            // Функция отображает список коробок с выключателями внутри в соответствии с выбранным номером заказа и версией в разделе 'Приход -> История'
            //
            List<string[]> list = new List<string[]>();
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT box_numb, part_code, part_name, amount, net, gross, black_id FROM items_in_boxes WHERE id_order='" + id_order + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string[] res_arr = new string[7];
                            res_arr[0] = reader.GetValue(0).ToString();
                            res_arr[1] = reader.GetValue(1).ToString();
                            res_arr[2] = reader.GetValue(2).ToString();
                            res_arr[3] = reader.GetValue(3).ToString();
                            res_arr[4] = reader.GetValue(4).ToString();
                            res_arr[5] = reader.GetValue(5).ToString();
                            res_arr[6] = reader.GetValue(6).ToString();
                            list.Add(res_arr);
                        }
                    }
                }
                connection.Close();
            }
            dataGridView8.RowCount = 0;
            dataGridView8.RowHeadersWidth = 35;
            dataGridView8.ColumnCount = 7;
            dataGridView8.RowCount = list.Count + 1;

            for (int i = 0; i < list.Count; i++)
            {
                dataGridView8.Rows[i].Cells[0].Value = list[i][0];
                dataGridView8.Rows[i].Cells[1].Value = list[i][1];
                dataGridView8.Rows[i].Cells[2].Value = list[i][2];
                dataGridView8.Rows[i].Cells[3].Value = list[i][3];
                dataGridView8.Rows[i].Cells[4].Value = list[i][4];
                dataGridView8.Rows[i].Cells[5].Value = list[i][5];
                dataGridView8.Rows[i].Cells[6].Value = list[i][6];
                if (i % 2 == 0)
                {
                    dataGridView8.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView8.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView8.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView8.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView8.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView8.Rows[i].Cells[5].Style.BackColor = cl_even;
                    dataGridView8.Rows[i].Cells[6].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView8.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView8.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView8.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView8.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView8.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView8.Rows[i].Cells[5].Style.BackColor = cl_odd;
                    dataGridView8.Rows[i].Cells[6].Style.BackColor = cl_odd;
                }
            }

            int indexOfTotalRow = 0;
            indexOfTotalRow = dataGridView8.RowCount - 1;
            int qty = 0;
            double netWeight = 0;
            double grossWeight = 0;
            for (int k = 0; k < dataGridView8.RowCount - 1; k++)
            {
                qty = qty + Convert.ToInt32(dataGridView8.Rows[k].Cells[3].Value);
                netWeight = netWeight + Math.Round(Convert.ToDouble(dataGridView8.Rows[k].Cells[4].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                grossWeight = grossWeight + Math.Round(Convert.ToDouble(dataGridView8.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
        }

            dataGridView8.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
            dataGridView8.Rows[indexOfTotalRow].Cells[3].Value = qty;
            dataGridView8.Rows[indexOfTotalRow].Cells[4].Value = netWeight;
            dataGridView8.Rows[indexOfTotalRow].Cells[5].Value = grossWeight;
            dataGridView8.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
        }

        public void LoadOrderItems(string id_order, int version)
        {
            // Функция отображает список выключателей в соответствии с выбранным номером заказа и версией в разделе 'История'
            //
            List<string[]> list = new List<string[]>();
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT id_item, part_code, name, amount, remarks FROM orderItems WHERE id_order='" + id_order + "' and version=" + version, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string[] res_arr = new string[5];
                            res_arr[0] = reader.GetValue(0).ToString();
                            res_arr[1] = reader.GetValue(1).ToString();
                            res_arr[2] = reader.GetValue(2).ToString();
                            res_arr[3] = reader.GetValue(3).ToString();
                            res_arr[4] = reader.GetValue(4).ToString();
                            list.Add(res_arr);
                        }
                    }
                }
                connection.Close();
            }
            //
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT remarks FROM orders WHERE id_order='" + id_order + "' and version=" + version, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            richTextBox2.Text = reader.GetValue(0).ToString();
                        }
                    }
                }
                connection.Close();
            }
            dataGridView6.RowCount = 0;
            dataGridView6.RowHeadersWidth = 35;
            //
            dataGridView6.ColumnCount = 8;
            dataGridView6.RowCount = list.Count;
            //
            dataGridView6.Columns[0].HeaderText = "#";
            dataGridView6.Columns[1].HeaderText = "Код продукта";
            dataGridView6.Columns[2].HeaderText = "Название продукта";
            dataGridView6.Columns[3].HeaderText = "Кол-во";
            dataGridView6.Columns[4].HeaderText = "";
            dataGridView6.Columns[5].HeaderText = "Не хватает";
            dataGridView6.Columns[6].HeaderText = "На складе";
            dataGridView6.Columns[7].HeaderText = "Заметки";
            //
            dataGridView6.Columns[0].Width = 50;
            dataGridView6.Columns[1].Width = 130;
            dataGridView6.Columns[2].Width = 200;
            dataGridView6.Columns[3].Width = 60;
            dataGridView6.Columns[4].Width = 30;
            dataGridView6.Columns[5].Width = 80;
            dataGridView6.Columns[6].Width = 80;
            int sz = dataGridView6.Columns[0].Width + dataGridView6.Columns[1].Width + dataGridView6.Columns[2].Width + dataGridView6.Columns[3].Width + dataGridView6.Columns[4].Width + dataGridView6.Columns[5].Width + dataGridView6.Columns[6].Width;
            dataGridView6.Columns[7].Width = dataGridView6.Size.Width - sz - 60;
            //
            for (int i = 0; i < list.Count; i++)
            {
                dataGridView6.Rows[i].Cells[0].Value = list[i][0];
                dataGridView6.Rows[i].Cells[1].Value = list[i][1];
                dataGridView6.Rows[i].Cells[2].Value = list[i][2];
                dataGridView6.Rows[i].Cells[3].Value = list[i][3];
                dataGridView6.Rows[i].Cells[7].Value = list[i][4];

                if (i % 2 == 0)
                {
                    dataGridView6.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[5].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[6].Style.BackColor = cl_even;
                    dataGridView6.Rows[i].Cells[7].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView6.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[5].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[6].Style.BackColor = cl_odd;
                    dataGridView6.Rows[i].Cells[7].Style.BackColor = cl_odd;
                }
            }

            int indexOfTotalRow = dataGridView6.RowCount;
            dataGridView6.RowCount = indexOfTotalRow + 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[k].Cells[3].Value);
            }
            dataGridView6.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
            dataGridView6.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
            dataGridView6.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
        }

        public void SelectOrder_NewVersion()
        {
            // Функция загружает в datagridview заказ из draft файла в разделе 'История -> создание новой версии'
            //
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(openFileDialog.FileName);
                ExcelPackage excel = new ExcelPackage(file);
                var worksheet = excel.Workbook.Worksheets[1];

                dataGridView6.RowCount = worksheet.Dimension.Rows - 1;
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView6.Rows[k].Cells[0].Value = worksheet.Cells[k + 2, 1].Value;
                    dataGridView6.Rows[k].Cells[1].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView6.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView6.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 4].Value;
                }
                
                //dataGridView1.Rows[7].Cells[2].Style.BackColor = Color.LightGreen;
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    int k = dataGridView6.Rows.IndexOf(row);
                    if (k % 2 == 0)
                    {
                        row.DefaultCellStyle.BackColor = cl_even;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = cl_odd;
                    }
                }

                dataGridView6.Columns[0].HeaderText = "#";
                dataGridView6.Columns[1].HeaderText = "Код продукта";
                dataGridView6.Columns[2].HeaderText = "Название продукта";
                dataGridView6.Columns[3].HeaderText = "Кол-во";
                dataGridView6.Columns[4].HeaderText = "";
                dataGridView6.Columns[5].HeaderText = "Не хватает";
                dataGridView6.Columns[6].HeaderText = "На складе";
                dataGridView6.Columns[7].HeaderText = "Заметки";
                //
                dataGridView6.Columns[0].Width = 50;
                dataGridView6.Columns[1].Width = 130;
                dataGridView6.Columns[2].Width = 200;
                dataGridView6.Columns[3].Width = 60;
                dataGridView6.Columns[3].Width = 30;
                dataGridView6.Columns[3].Width = 60;
                dataGridView6.Columns[3].Width = 60;
                int sz = dataGridView6.Columns[0].Width + dataGridView6.Columns[1].Width + dataGridView6.Columns[2].Width + dataGridView6.Columns[3].Width + dataGridView6.Columns[4].Width + dataGridView6.Columns[5].Width + dataGridView6.Columns[6].Width;
                dataGridView6.Columns[7].Width = dataGridView6.Size.Width - sz - 60;

                int indexOfTotalRow = dataGridView6.RowCount;
                dataGridView6.RowCount = indexOfTotalRow + 1;
                int totalQty = 0;
                for (int k = 0; k < indexOfTotalRow; k++)
                {
                    totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[k].Cells[3].Value);
                }
                dataGridView6.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
                dataGridView6.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
                dataGridView6.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;

                button11_save_v.Enabled = true;
            }
        }

        public void SelectOrder()
        {
            // Функция загружает в datagridview заказ из draft файла в разделе 'Создание заказа'
            //
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(openFileDialog.FileName);
                ExcelPackage excel = new ExcelPackage(file);
                var worksheet = excel.Workbook.Worksheets[1];

                dataGridView1.RowCount = worksheet.Dimension.Rows - 1;
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView1.Rows[k].Cells[0].Value = worksheet.Cells[k + 2, 1].Value;
                    dataGridView1.Rows[k].Cells[1].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView1.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView1.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 4].Value;
                    dataGridView1.Rows[k].Cells[4].Value = worksheet.Cells[k + 2, 5].Value;
                    dataGridView1.Rows[k].HeaderCell.Value = "";
                }

                int i = 0;
                for (i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (dataGridView1.Columns[i].HeaderText.ToString().Length < 0)
                    {
                        dataGridView1.Columns.RemoveAt(i);
                    }
                }
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    int k = dataGridView1.Rows.IndexOf(row);
                    if (k % 2 == 0)
                    {
                        row.DefaultCellStyle.BackColor = cl_even;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = cl_odd;
                    }
                    row.HeaderCell.Value = "";
                }

                dataGridView1.RowHeadersWidth = 35;
            }
            openFileDialog.Dispose();

            int indexOfTotalRow = dataGridView1.RowCount;
            dataGridView1.RowCount = indexOfTotalRow + 1;
            int totalQty = 0;
            for(int k=0; k<indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[k].Cells[3].Value);
            }
            dataGridView1.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
            dataGridView1.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
            dataGridView1.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
        }

        //}
        
        public void UploadOrder_NewVersion() {
            // Функция загружает заказ со списком продуктов в базу данных из datagridview в разделе 'Создание заказа -> добавить новую версию'
            //
            int selected = -1;
            //
            //
            string year = Convert.ToString(dateTimePicker3.Value.Year);
            string month = Convert.ToString(dateTimePicker3.Value.Month);
            string day = Convert.ToString(dateTimePicker3.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string v_date = day + "." + month + "." + year;
            //
            selected = dataGridView4.CurrentCell.RowIndex;
            string id_order = dataGridView4.Rows[selected].Cells[2].Value.ToString();
            string dz_date = dataGridView4.Rows[selected].Cells[3].Value.ToString();
            string c_order = dataGridView4.Rows[selected].Cells[4].Value.ToString();
            string c_date = dataGridView4.Rows[selected].Cells[5].Value.ToString();
            string descr = dataGridView4.Rows[selected].Cells[0].Value.ToString();
            string cp = dataGridView4.Rows[selected].Cells[6].Value.ToString();
            string cp_date = dataGridView4.Rows[selected].Cells[7].Value.ToString();
            string ct = dataGridView4.Rows[selected].Cells[8].Value.ToString();
            string ct_date = dataGridView4.Rows[selected].Cells[9].Value.ToString();
            string remarks = richTextBox2.Text;
            int version = dataGridView5.RowCount + 1;

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "INSERT INTO orders(id_order, customer_order_id, dz_date, c_date, descr, version, com_proposal, contract, remarks, com_proposal_date, contract_date) VALUES(@orderid, @customerorderid, @dz_date, @c_date, @descr, @version, @cp, @ct, @remarks, @cp_date, @ct_date)";
            command.Parameters.AddWithValue("@orderid", id_order);
            command.Parameters.AddWithValue("@customerorderid", c_order);
            command.Parameters.AddWithValue("@dz_date", dz_date);
            command.Parameters.AddWithValue("@c_date", c_date);
            command.Parameters.AddWithValue("@descr", descr);
            command.Parameters.AddWithValue("@version", version);
            command.Parameters.AddWithValue("@cp", cp);
            command.Parameters.AddWithValue("@ct", ct);
            command.Parameters.AddWithValue("@remarks", remarks);
            command.Parameters.AddWithValue("@cp_date", cp_date);
            command.Parameters.AddWithValue("@ct_date", ct_date);
            command.ExecuteNonQuery();
            //
            int i = 0;
            string partCode;
            string name;
            string amount;
            string remark;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
            {
                string id_item = (i + 1).ToString();
                try { partCode = dataGridView6.Rows[i].Cells[1].Value.ToString(); } catch { partCode = ""; };
                try { name = dataGridView6.Rows[i].Cells[2].Value.ToString(); } catch { name = ""; };
                try { amount = dataGridView6.Rows[i].Cells[3].Value.ToString(); } catch { amount = ""; };
                try { remark = dataGridView6.Rows[i].Cells[7].Value.ToString(); } catch { remark = ""; };
                //
                SqlCommand command1 = new SqlCommand();
                command1.Connection = connection;
                command1.CommandType = CommandType.Text;
                command1.CommandText = "INSERT INTO orderItems(id_item, part_code, name, amount, id_order, version, version_date, remarks) VALUES(@id_item, @partcode, @name, @amount, @idorder, @version, @version_date, @remarks)";
                command1.Parameters.AddWithValue("@id_item", id_item);
                command1.Parameters.AddWithValue("@partcode", partCode);
                command1.Parameters.AddWithValue("@name", name);
                command1.Parameters.AddWithValue("@amount", amount);
                command1.Parameters.AddWithValue("@idorder", id_order);
                command1.Parameters.AddWithValue("@version", version);
                command1.Parameters.AddWithValue("@version_date", v_date);
                command1.Parameters.AddWithValue("@remarks", remark);
                command1.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("Новая версия успешно добавлена к выбранному заказу.", "Сообщение",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
            button8_search.Enabled = true;
            button9_add_v.Enabled = true;
            button16_edit_v.Enabled = true;
            button12_gen_doc.Enabled = true;
            button15_cancel.Enabled = false;
            dataGridView4.Enabled = true;
            dataGridView5.Enabled = true;
            dataGridView6.AllowUserToAddRows = false;
            dataGridView1.DataSource = null;
            menuStrip1.Items[0].Enabled = true;
            GetVersionsList(id_order);
        }

        public void UploadOrder() {
            // Функция загружает заказ со списком продуктов в базу данных из datagridview в разделе 'Создание заказа'
            //
            string id_item;
            string partCode;
            string name;
            string amount;
            string idOrder;
            string remarks;
            Guid guid = Guid.NewGuid();
            //idOrder = guid.ToString();
            DateTime dt = DateTime.Now;
            string s1 = dt.Day.ToString();
            string s2 = dt.Month.ToString();
            string s3 = dt.Year.ToString();
            s3 = s3.Substring(2, 2);
            string s4 = dt.Hour.ToString();
            string s5 = dt.Minute.ToString();
            string s6 = dt.Second.ToString();
            if (s1.Length == 1) { s1 = "0" + s1; }
            if (s2.Length == 1) { s2 = "0" + s2; }
            if (s4.Length == 1) { s4 = "0" + s4; }
            if (s5.Length == 1) { s5 = "0" + s5; }
            if (s6.Length == 1) { s6 = "0" + s6; }
            idOrder = s1 + s2 + s3 + s4 + s5 + s6;
            //
            DateTime dateTime = DateTime.Now;
            string year = dateTime.Year.ToString();
            string month = dateTime.Month.ToString();
            string day = dateTime.Day.ToString();
            if (month.Length ==1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string dz_date = day + "." + month + "." + year;
            //
            // Customer order date
            year = Convert.ToString(dateTimePicker1.Value.Year);
            month = Convert.ToString(dateTimePicker1.Value.Month);
            day = Convert.ToString(dateTimePicker1.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string c_date = day + "." + month + "." + year;
            //
            // Comercial proposal date
            year = Convert.ToString(dateTimePicker7.Value.Year);
            month = Convert.ToString(dateTimePicker7.Value.Month);
            day = Convert.ToString(dateTimePicker7.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string cp_date = day + "." + month + "." + year;
            //
            // Contract date
            year = Convert.ToString(dateTimePicker8.Value.Year);
            month = Convert.ToString(dateTimePicker8.Value.Month);
            day = Convert.ToString(dateTimePicker8.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string ct_date = day + "." + month + "." + year;
            //
            string customer_order_id;
            if (checkBox2.Checked == true)
            {
                if (textBox1.Text.Length > 0)
                {
                    textBox2.Text = idOrder;
                    customer_order_id = textBox1.Text.ToString();
                    SqlConnection connection = new SqlConnection(conString);
                    connection.Open();
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO orders(id_order, customer_order_id, dz_date, c_date, descr, com_proposal, contract, remarks, com_proposal_date, contract_date) VALUES(@orderid, @customerorderid, @dz_date, @c_date, @descr, @cp, @ct, @remarks, @cp_date, @ct_date)";
                    command.Parameters.AddWithValue("@orderid", idOrder);
                    command.Parameters.AddWithValue("@customerorderid", customer_order_id);
                    command.Parameters.AddWithValue("@dz_date", dz_date);
                    command.Parameters.AddWithValue("@c_date", c_date);
                    command.Parameters.AddWithValue("@descr", textBox13.Text);
                    command.Parameters.AddWithValue("@cp", textBox5.Text);
                    command.Parameters.AddWithValue("@ct", textBox6.Text);
                    command.Parameters.AddWithValue("@remarks", richTextBox1.Text);
                    command.Parameters.AddWithValue("@cp_date", cp_date);
                    command.Parameters.AddWithValue("@ct_date", ct_date);
                    command.ExecuteNonQuery();
                    //
                    int i = 0;
                    for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        id_item = (i + 1).ToString();
                        try { partCode = dataGridView1.Rows[i].Cells[1].Value.ToString(); } catch { partCode = ""; };
                        try { name = dataGridView1.Rows[i].Cells[2].Value.ToString(); } catch { name = ""; };
                        try { amount = dataGridView1.Rows[i].Cells[3].Value.ToString(); } catch { amount = ""; };
                        try { remarks = dataGridView1.Rows[i].Cells[7].Value.ToString(); } catch { remarks = ""; };
                        SqlCommand command1 = new SqlCommand();
                        command1.Connection = connection;
                        command1.CommandType = CommandType.Text;
                        command1.CommandText = "INSERT INTO orderItems(id_item, part_code, name, amount, id_order, version_date, remarks) VALUES(@id_item, @partcode, @name, @amount, @idorder, @version_date, @remarks)";
                        command1.Parameters.AddWithValue("@partcode", partCode);
                        command1.Parameters.AddWithValue("@name", name);
                        command1.Parameters.AddWithValue("@amount", amount);
                        command1.Parameters.AddWithValue("@idorder", idOrder);
                        command1.Parameters.AddWithValue("@id_item", id_item);
                        command1.Parameters.AddWithValue("@version_date", dz_date);
                        command1.Parameters.AddWithValue("@remarks", remarks);
                        command1.ExecuteNonQuery();
                    }
                    // Update com_proposal table
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "UPDATE com_proposal SET po_id=@po, po_date=@po_dt WHERE cp_id=@cp and version=@v";
                    command.Parameters.AddWithValue("@po", idOrder);
                    command.Parameters.AddWithValue("@po_dt", dz_date);
                    command.Parameters.AddWithValue("@cp", textBox5.Text);
                    command.Parameters.AddWithValue("@v", label104.Text);
                    command.ExecuteNonQuery();
                    // Update contracts table
                    command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "UPDATE contracts SET id_order=@po, id_order_date=@po_dt WHERE id_cp=@cp and cp_version=@v";
                    command.Parameters.AddWithValue("@po", idOrder);
                    command.Parameters.AddWithValue("@po_dt", dz_date);
                    command.Parameters.AddWithValue("@cp", textBox5.Text);
                    command.Parameters.AddWithValue("@v", label104.Text);
                    command.ExecuteNonQuery();
                    //
                    connection.Close();
                    //
                    MessageBox.Show("Заказ успешно загружен в систему", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    button2.Enabled = false;
                    button1.Enabled = false;
                    button3.Enabled = true;
                    button14.Enabled = true;
                    button13.Enabled = false;
                    textBox1.ReadOnly = true;
                    textBox5.ReadOnly = true;
                    textBox6.ReadOnly = true;
                    dateTimePicker1.Enabled = false;
                }
                else
                {
                    MessageBox.Show("Необходимо указать номер заказа от клиента", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                textBox2.Text = idOrder;
                customer_order_id = "";
                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "INSERT INTO orders(id_order, customer_order_id, dz_date, c_date, descr, com_proposal, contract, remarks, com_proposal_date, contract_date) VALUES(@orderid, @customerorderid, @dz_date, @c_date, @descr, @cp, @ct, @remarks, @cp_date, @ct_date)";
                command.Parameters.AddWithValue("@orderid", idOrder);
                command.Parameters.AddWithValue("@customerorderid", customer_order_id);
                command.Parameters.AddWithValue("@dz_date", dz_date);
                command.Parameters.AddWithValue("@c_date", "");
                command.Parameters.AddWithValue("@descr", "Drone Zone");
                command.Parameters.AddWithValue("@cp", textBox5.Text);
                command.Parameters.AddWithValue("@ct", textBox6.Text);
                command.Parameters.AddWithValue("@remarks", richTextBox1.Text);
                command.Parameters.AddWithValue("@cp_date", "");
                command.Parameters.AddWithValue("@ct_date", "");
                command.ExecuteNonQuery();
                //
                int i = 0;
                for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    id_item = (i + 1).ToString();
                    try { partCode = dataGridView1.Rows[i].Cells[1].Value.ToString(); } catch { partCode = ""; };
                    try { name = dataGridView1.Rows[i].Cells[2].Value.ToString(); } catch { name = ""; };
                    try { amount = dataGridView1.Rows[i].Cells[3].Value.ToString(); } catch { amount = ""; };
                    try { remarks = dataGridView1.Rows[i].Cells[7].Value.ToString(); } catch { remarks = ""; };
                    SqlCommand command1 = new SqlCommand();
                    command1.Connection = connection;
                    command1.CommandType = CommandType.Text;
                    command1.CommandText = "INSERT INTO orderItems(id_item, part_code, name, amount, id_order, version_date, remarks) VALUES(@id_item, @partcode, @name, @amount, @idorder, @version_date, @remarks)";
                    command1.Parameters.AddWithValue("@partcode", partCode);
                    command1.Parameters.AddWithValue("@name", name);
                    command1.Parameters.AddWithValue("@amount", amount);
                    command1.Parameters.AddWithValue("@idorder", idOrder);
                    command1.Parameters.AddWithValue("@id_item", id_item);
                    command1.Parameters.AddWithValue("@version_date", dz_date);
                    command1.Parameters.AddWithValue("@remarks", remarks);
                    command1.ExecuteNonQuery();
                }
                // Update com_proposal table
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE com_proposal SET po_id=@po, po_date=@po_dt WHERE cp_id=@cp and version=@v";
                command.Parameters.AddWithValue("@po", idOrder);
                command.Parameters.AddWithValue("@po_dt", dz_date);
                command.Parameters.AddWithValue("@cp", textBox5.Text);
                command.Parameters.AddWithValue("@v", label104.Text);
                command.ExecuteNonQuery();
                // Update contracts table
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "UPDATE contracts SET id_order=@po, id_order_date=@po_dt WHERE cp_id=@cp and version=@v";
                command.Parameters.AddWithValue("@po", idOrder);
                command.Parameters.AddWithValue("@po_dt", dz_date);
                command.Parameters.AddWithValue("@cp", textBox5.Text);
                command.Parameters.AddWithValue("@v", label104.Text);
                command.ExecuteNonQuery();
                //
                connection.Close();
                MessageBox.Show("Заказ успешно загружен в систему", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
                button1.Enabled = false;
                button3.Enabled = true;
                button14.Enabled = true;
                button13.Enabled = false;
                textBox1.ReadOnly = true;
                textBox5.ReadOnly = true;
                textBox6.ReadOnly = true;
                dateTimePicker1.Enabled = false;
            }
            dataGridView1.DataSource = null;
        }

        public void GenerateOrderHist(string po, string po_date, string order, string order_date, string cp, string cp_date, string ct, string ct_date, string v, string company)
        {

            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");
            worksheet.DefaultColWidth = 10;
            worksheet.Column(1).Width = 5;
            worksheet.Column(2).Width = 16;
            worksheet.Column(8).Width = 8;
            worksheet.Column(9).Width = 18;
            using (System.Drawing.Image logo = System.Drawing.Image.FromFile(app_dir_temp + "Logo.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("My Logo", logo);
                excelImage.SetPosition(1, 0, 2, 100);
                excelImage.SetSize(25);
            }
            worksheet.Cells[1, 8].Value = "PRODUCTION";
            worksheet.Cells[1, 8].Style.Font.Size = 14;
            worksheet.Cells[1, 8, 1, 9].Merge = true;
            worksheet.Cells[2, 8].Value = "WORK ORDER";
            worksheet.Cells[2, 8].Style.Font.Size = 14;
            worksheet.Cells[2, 8, 2, 9].Merge = true;

            worksheet.Cells[9, 1].Value = "DRONE ZONE CO, LTD";
            worksheet.Cells[9, 1, 9, 3].Merge = true;
            worksheet.Cells[10, 1].Value = "30F B-Dong, 323 Incheon Tower Daero,";
            worksheet.Cells[10, 1, 10, 3].Merge = true;
            worksheet.Cells[11, 1].Value = "Yeonsu-Gu, Incheon, Republic of Korea";
            worksheet.Cells[11, 1, 11, 3].Merge = true;
            worksheet.Cells[12, 1].Value = "Tel: +82-10-2206-8299";
            worksheet.Cells[12, 1, 12, 3].Merge = true;
            worksheet.Cells[13, 1].Value = "E-mail: dronezone.anna@gmail.com";
            worksheet.Cells[13, 1, 13, 3].Merge = true;

            worksheet.Cells[16, 1].Value = "TO: LSIS CO., LTD.";
            worksheet.Cells[16, 1, 16, 3].Merge = true;
            worksheet.Cells[17, 1].Value = "LS Tower 1026-6, Hogye-dong, Dongan-gu,";
            worksheet.Cells[17, 1, 17, 3].Merge = true;
            worksheet.Cells[18, 1].Value = "Anyang-si, Gyeonggi-do 431-848,";
            worksheet.Cells[18, 1, 18, 3].Merge = true;
            worksheet.Cells[19, 1].Value = "Seoul, Republic of Korea";
            worksheet.Cells[19, 1, 19, 3].Merge = true;
            worksheet.Cells[20, 1].Value = "Tel.: +82-2-2034-4429";
            worksheet.Cells[20, 1, 20, 3].Merge = true;
            worksheet.Cells[21, 1].Value = "E-mail: skkanga@lsis.com";
            worksheet.Cells[21, 1, 21, 3].Merge = true;

            worksheet.Cells[16, 4].Value = "TO: " + company;
            worksheet.Cells[16, 4, 16, 6].Merge = true;
            worksheet.Cells[17, 4].Value = "ORDER:       #" + order + "  Date: " + order_date;
            worksheet.Cells[17, 4, 17, 6].Merge = true;
            worksheet.Cells[18, 4].Value = "SHIP IN KOREA TO: Incheon Port ";
            worksheet.Cells[18, 4, 18, 6].Merge = true;
            worksheet.Cells[19, 4].Value = "FINAL DESTINATION:";
            worksheet.Cells[19, 4, 19, 6].Merge = true;

            worksheet.Cells[16, 7].Value = "PO: " + po + "-" + v + "  Date: " + po_date;
            worksheet.Cells[16, 7, 16, 9].Merge = true;
            worksheet.Cells[17, 7].Value = "CP: " + cp + "  Date: " + cp_date;
            worksheet.Cells[17, 7, 17, 9].Merge = true;
            worksheet.Cells[18, 7].Value = "CT: " + ct + "  Date:" + ct_date;
            worksheet.Cells[18, 7, 18, 9].Merge = true;

            worksheet.Cells[26, 1].Value = "    Drone Zone Co., LTD would like to thank you for your business and request to start production of ";
            worksheet.Cells[26, 1, 26, 9].Merge = true;
            worksheet.Cells[27, 1].Value = "    following components. Please issue an invoice,  stating the amount to be paid for this order. ";
            worksheet.Cells[27, 1, 27, 9].Merge = true;

            worksheet.Cells[29, 1].Value = "#";
            worksheet.Cells[29, 2].Value = "Item Code";
            worksheet.Cells[29, 3].Value = "Item Name";
            worksheet.Cells[29, 3, 29, 5].Merge = true;
            worksheet.Cells[29, 6].Value = "QTY";
            worksheet.Cells[29, 7].Value = "Notes";
            worksheet.Cells[29, 7, 29, 9].Merge = true;
            using (ExcelRange rng = worksheet.Cells[29, 1, 29, worksheet.Dimension.Columns])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }

            int rowCount = 30;
            int totalQty = 0;

            for (int i = 0; i < dataGridView6.RowCount - 1; i++)
            {
                worksheet.Cells[rowCount, 1].Value = dataGridView6.Rows[i].Cells[0].Value;
                worksheet.Cells[rowCount, 2].Value = dataGridView6.Rows[i].Cells[1].Value;
                worksheet.Cells[rowCount, 3].Value = dataGridView6.Rows[i].Cells[2].Value;
                worksheet.Cells[rowCount, 3, rowCount, 5].Merge = true;
                worksheet.Cells[rowCount, 6].Value = dataGridView6.Rows[i].Cells[3].Value;
                worksheet.Cells[rowCount, 7].Value = dataGridView6.Rows[i].Cells[7].Value;
                worksheet.Cells[rowCount, 7, rowCount, 9].Merge = true;

                totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[i].Cells[3].Value);
                rowCount++;
            }
            worksheet.Cells[rowCount, 1].Value = "TOTAL:";
            worksheet.Cells[rowCount, 1, rowCount, 5].Merge = true;
            worksheet.Cells[rowCount, 6].Value = totalQty;
            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange rng = worksheet.Cells[29, 1, rowCount, 9])
            {
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[30, 1, rowCount-1, 9])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }

            rowCount++;
            for (int k = 0; k < richTextBox1.Lines.Count(); k++)
            {
                worksheet.Cells[rowCount, 1].Value = richTextBox1.Lines[k].ToString();
                worksheet.Cells[rowCount, 1, rowCount, 9].Merge = true;
                rowCount++;
            }

            using (System.Drawing.Image sign = System.Drawing.Image.FromFile(app_dir_temp + "Sign.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Sign", sign);
                excelImage.SetPosition(rowCount + 1, 15, 5, 0);
                excelImage.SetSize(14);
            }
            using (System.Drawing.Image stamp = System.Drawing.Image.FromFile(app_dir_temp + "Stamp.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Stamp", stamp);
                excelImage.SetPosition(rowCount + 1, 0, 7, 15);
                excelImage.SetSize(20);
            }
            worksheet.Cells[rowCount + 4, 4, rowCount + 4, 9].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[rowCount + 5, 4].Value = " Authorized and signed by Anna Li";
            worksheet.Cells[rowCount + 5, 4, rowCount + 5, 9].Merge = true;
            worksheet.Cells[rowCount + 6, 4].Value = " by Power of Attorney on behalf of ";
            worksheet.Cells[rowCount + 6, 4, rowCount + 6, 9].Merge = true;
            worksheet.Cells[rowCount + 7, 4].Value = " Drone Zone General Manager  ";
            worksheet.Cells[rowCount + 7, 4, rowCount + 7, 9].Merge = true;
            worksheet.Cells[rowCount + 8, 4].Value = " dronezone.anna@gmail.com  ";
            worksheet.Cells[rowCount + 8, 4].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 4, rowCount + 8, 9].Merge = true;


            worksheet.Cells[rowCount + 7, 1].Value = " Customer care  ";
            worksheet.Cells[rowCount + 7, 1, rowCount + 7, 3].Merge = true;
            worksheet.Cells[rowCount + 8, 1].Value = " dronezone.sk@gmail.com  ";
            worksheet.Cells[rowCount + 8, 1].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 1, rowCount + 8, 3].Merge = true;

            using (ExcelRange rng = worksheet.Cells[9, 1, rowCount + 10, worksheet.Dimension.Columns])
            {
                rng.Style.Font.Size = 8;
                rng.Style.Font.Name = "Times New Roman";
            }

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                excel.SaveAs(new FileInfo(saveFileDialog.FileName));
            }
        }
        public void GenerateOrder()
        {
            DateTime dateTime = DateTime.Now;
            string year = dateTime.Year.ToString();
            string month = dateTime.Month.ToString();
            string day = dateTime.Day.ToString();
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string dz_date = day + "." + month + "." + year;

            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");
            worksheet.DefaultColWidth = 10;
            worksheet.Column(1).Width = 5;
            worksheet.Column(2).Width = 16;
            worksheet.Column(8).Width = 8;
            worksheet.Column(9).Width = 18;
            using (System.Drawing.Image logo = System.Drawing.Image.FromFile(app_dir_temp + "Logo.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("My Logo", logo);
                excelImage.SetPosition(1, 0, 2, 100);
                excelImage.SetSize(25);
            }
            worksheet.Cells[1, 8].Value = "PRODUCTION";
            worksheet.Cells[1, 8].Style.Font.Size = 14;
            worksheet.Cells[1, 8, 1, 9].Merge = true;
            worksheet.Cells[2, 8].Value = "WORK ORDER";
            worksheet.Cells[2, 8].Style.Font.Size = 14;
            worksheet.Cells[2, 8, 2, 9].Merge = true;

            worksheet.Cells[9, 1].Value = "DRONE ZONE CO, LTD";
            worksheet.Cells[9, 1, 9, 3].Merge = true;
            worksheet.Cells[10, 1].Value = "30F B-Dong, 323 Incheon Tower Daero,";
            worksheet.Cells[10, 1, 10, 3].Merge = true;
            worksheet.Cells[11, 1].Value = "Yeonsu-Gu, Incheon, Republic of Korea";
            worksheet.Cells[11, 1, 11, 3].Merge = true;
            worksheet.Cells[12, 1].Value = "Tel: +82-10-2206-8299";
            worksheet.Cells[12, 1, 12, 3].Merge = true;
            worksheet.Cells[13, 1].Value = "E-mail: dronezone.anna@gmail.com";
            worksheet.Cells[13, 1, 13, 3].Merge = true;

            worksheet.Cells[16, 1].Value = "TO: LS Electric CO., LTD.";
            worksheet.Cells[16, 1, 16, 3].Merge = true;
            worksheet.Cells[17, 1].Value = "LS Tower 1026-6, Hogye-dong, Dongan-gu,";
            worksheet.Cells[17, 1, 17, 3].Merge = true;
            worksheet.Cells[18, 1].Value = "Anyang-si, Gyeonggi-do 431-848,";
            worksheet.Cells[18, 1, 18, 3].Merge = true;
            worksheet.Cells[19, 1].Value = "Seoul, Republic of Korea";
            worksheet.Cells[19, 1, 19, 3].Merge = true;
            worksheet.Cells[20, 1].Value = "Tel.: +82-2-2034-4429";
            worksheet.Cells[20, 1, 20, 3].Merge = true;
            worksheet.Cells[21, 1].Value = "E-mail: skkanga@lsis.com";
            worksheet.Cells[21, 1, 21, 3].Merge = true;

            worksheet.Cells[16, 4].Value = "TO: " + textBox13.Text;
            worksheet.Cells[16, 4, 16, 6].Merge = true;
            worksheet.Cells[17, 4].Value = "ORDER:       #" + textBox1.Text + "  Date: " + dateTimePicker1.Value.Date.ToString("dd.MM.yyyy");
            worksheet.Cells[17, 4, 17, 6].Merge = true;
            worksheet.Cells[18, 4].Value = "Project: " + textBox36.Text;
            worksheet.Cells[18, 4, 18, 6].Merge = true;
            worksheet.Cells[19, 4].Value = "SHIP BY: AIR / RW";
            worksheet.Cells[19, 4, 19, 6].Merge = true;
            worksheet.Cells[20, 4].Value = "FINAL DESTINATION:";
            worksheet.Cells[20, 4, 20, 6].Merge = true;

            worksheet.Cells[16, 7].Value = "PO: " + textBox2.Text + "  Date: " + dz_date;
            worksheet.Cells[16, 7, 16, 9].Merge = true;
            worksheet.Cells[17, 7].Value = "CP: " + textBox5.Text + "  Date: " + dateTimePicker7.Value.Date.ToString("dd.MM.yyyy");
            worksheet.Cells[17, 7, 17, 9].Merge = true;
            worksheet.Cells[18, 7].Value = "CT: " + textBox6.Text + "  Date:" + dateTimePicker8.Value.Date.ToString("dd.MM.yyyy");
            worksheet.Cells[18, 7, 18, 9].Merge = true;

            worksheet.Cells[26, 1].Value = "    Drone Zone Co., LTD would like to thank you for your business and request to start production of ";
            worksheet.Cells[26, 1, 26, 9].Merge = true;
            worksheet.Cells[27, 1].Value = "    following components. Please issue an invoice,  stating the amount to be paid for this order. ";
            worksheet.Cells[27, 1, 27, 9].Merge = true;

            worksheet.Cells[29, 1].Value = "#";
            worksheet.Cells[29, 2].Value = "Item Code";
            worksheet.Cells[29, 3].Value = "Item Name";
            worksheet.Cells[29, 3, 29, 5].Merge = true;
            worksheet.Cells[29, 6].Value = "QTY";
            worksheet.Cells[29, 7].Value = "Notes";
            worksheet.Cells[29, 7, 29, 9].Merge = true;
            using (ExcelRange rng = worksheet.Cells[29, 1, 29, worksheet.Dimension.Columns])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }

            int rowCount = 30;
            int totalQty = 0;
            
            for(int i = 0; i<dataGridView1.RowCount-1; i++)
            {
                worksheet.Cells[rowCount, 1].Value = dataGridView1.Rows[i].Cells[0].Value;
                worksheet.Cells[rowCount, 2].Value = dataGridView1.Rows[i].Cells[1].Value;
                worksheet.Cells[rowCount, 3].Value = dataGridView1.Rows[i].Cells[2].Value;
                worksheet.Cells[rowCount, 3, rowCount, 5].Merge = true;
                worksheet.Cells[rowCount, 6].Value = dataGridView1.Rows[i].Cells[3].Value;
                worksheet.Cells[rowCount, 7].Value = dataGridView1.Rows[i].Cells[7].Value;
                worksheet.Cells[rowCount, 7, rowCount, 9].Merge = true;

                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);
                rowCount++;
            }
            worksheet.Cells[rowCount, 1].Value = "TOTAL:";
            worksheet.Cells[rowCount, 1, rowCount, 5].Merge = true;
            worksheet.Cells[rowCount, 6].Value = totalQty;
            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 9])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange rng = worksheet.Cells[29, 1, rowCount, 9])
            {
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[30, 1, rowCount-1, 9])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }

            rowCount++;
            for (int k = 0; k < richTextBox1.Lines.Count(); k++)
            {
                worksheet.Cells[rowCount, 1].Value = richTextBox1.Lines[k].ToString();
                worksheet.Cells[rowCount, 1, rowCount, 9].Merge = true;
                rowCount++;
            }

            using (System.Drawing.Image sign = System.Drawing.Image.FromFile(app_dir_temp + "Sign.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Sign", sign);
                excelImage.SetPosition(rowCount + 1, 15, 5, 0);
                excelImage.SetSize(14);
            }
            using (System.Drawing.Image stamp = System.Drawing.Image.FromFile(app_dir_temp + "Stamp.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Stamp", stamp);
                excelImage.SetPosition(rowCount + 1, 0, 7, 15);
                excelImage.SetSize(20);
            }
            worksheet.Cells[rowCount + 4, 4, rowCount + 4, 9].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[rowCount + 5, 4].Value = " Authorized and signed by Anna Li";
            worksheet.Cells[rowCount + 5, 4, rowCount + 5, 9].Merge = true;
            worksheet.Cells[rowCount + 6, 4].Value = " by Power of Attorney on behalf of ";
            worksheet.Cells[rowCount + 6, 4, rowCount + 6, 9].Merge = true;
            worksheet.Cells[rowCount + 7, 4].Value = " Drone Zone General Manager  ";
            worksheet.Cells[rowCount + 7, 4, rowCount + 7, 9].Merge = true;
            worksheet.Cells[rowCount + 8, 4].Value = " dronezone.anna@gmail.com  ";
            worksheet.Cells[rowCount + 8, 4].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 4, rowCount + 8, 9].Merge = true;


            worksheet.Cells[rowCount + 7, 1].Value = " Customer care  ";
            worksheet.Cells[rowCount + 7, 1, rowCount + 7, 3].Merge = true;
            worksheet.Cells[rowCount + 8, 1].Value = " dronezone.sk@gmail.com  ";
            worksheet.Cells[rowCount + 8, 1].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 8, 1, rowCount + 8, 3].Merge = true;

            using (ExcelRange rng = worksheet.Cells[9, 1, rowCount + 10, worksheet.Dimension.Columns])
            {
                rng.Style.Font.Size = 8;
                rng.Style.Font.Name = "Times New Roman";
            }

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                excel.SaveAs(new FileInfo(saveFileDialog.FileName));
            }
        }

        List<string> GetShelfsStruct()
        {
            // Выборка из базы данных структуры стелажей на складе то есть кол-во стелажей, ярусов и секций в разделе 'Струкрура склада'
            //
            List<string> list = new List<string>();
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT * FROM Shelfs", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(reader.GetValue(0).ToString());
                            list.Add(reader.GetValue(1).ToString());
                            list.Add(reader.GetValue(2).ToString());
                        }
                    }
                }
                connection.Close();
            }
            return list;
        }

        public void UpdateShelfsStruct(int z, int x, int y)
        {
            // Запись в базу данных структуры стелажей на складе то есть кол-во стелажей, ярусов и секций в разделе 'Струкрура склада'
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "DELETE FROM Shelfs";
            command.ExecuteNonQuery();
            connection.Close();
            //
            connection = new SqlConnection(conString);
            connection.Open();
            command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "INSERT INTO Shelfs(z, x, y) VALUES(@z, @x, @y)";
            command.Parameters.AddWithValue("@z", z);
            command.Parameters.AddWithValue("@x", x);
            command.Parameters.AddWithValue("@y", y);
            command.ExecuteNonQuery();
            connection.Close();
        }

        bool shelfs_list_builded = false;
        public void MakeShelfsStructure(int z) {
            // Отображение структуры склада в разделе 'Струкрура склада'
            //
            List<string> shelf_struct = new List<string>();
            //List<BoxLocation> box_locations = new List<BoxLocation>();
            //
            shelf_struct = GetShelfsStruct();
            //box_locations = GetBoxesLocations(3);
            if (shelf_struct.Count != 0 && shelf_struct[0] != "0")
            {
                label11.Text = "0";
                // стелажи (ось Z)
                //string s = numericUpDown1.Value.ToString();
                string s;
                if (shelfs_list_builded == false) {
                    shelfs_list_builded = true;
                    checkedListBox1.Items.Clear();
                    s = shelf_struct[0];
                    numericUpDown1.Value = Convert.ToDecimal(s);
                    for (int i = 0; i < Int32.Parse(s); i++) {
                        checkedListBox1.Items.Add("Стелаж " + (i + 1));
                    }
                }
                // секции (ось X)
                //s = numericUpDown2.Value.ToString();
                dataGridView2.RowCount = 0;
                s = shelf_struct[1];
                numericUpDown2.Value = Convert.ToDecimal(s);
                dataGridView2.ColumnCount = Int32.Parse(s);
                int grid_w = (dataGridView2.Width - 100) / Int32.Parse(s);
                for (int i = 0; i < Int32.Parse(s); i++)
                {
                    dataGridView2.Columns[i].HeaderText = "Секция" + (i + 1);
                    if (grid_w < 60)
                    {
                        dataGridView2.Columns[i].Width = 70;
                    }
                    else
                    {
                        dataGridView2.Columns[i].Width = grid_w;
                    }
                }
                dataGridView2.RowHeadersWidth = 85;
                // секции (ось Y)
                //s = numericUpDown3.Value.ToString();
                s = shelf_struct[2];
                numericUpDown3.Value = Convert.ToDecimal(s);
                dataGridView2.RowCount = Int32.Parse(s);
                int grid_h = (dataGridView2.Height - 30) / Int32.Parse(s);
                for (int i = 0; i < Int32.Parse(s); i++)
                {
                    dataGridView2.Rows[i].HeaderCell.Value = "Ярус" + (Int32.Parse(s) - i);
                    if (grid_h < 20)
                    {
                        dataGridView2.Rows[i].Height = 22;
                    }
                    else
                    {
                        dataGridView2.Rows[i].Height = grid_h;
                    }
                }
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    for (int n = 0; n < dataGridView2.ColumnCount; n++)
                    {
                        dataGridView2.Rows[dataGridView2.RowCount - (i + 1)].Cells[n].Value = (i + 1) + "." + (n + 1);
                        // Mark empty shelf
                        DataGridViewImageCell iCell = new DataGridViewImageCell();
                        System.Drawing.Image img = Properties.Resources.shelf_empty2;
                        iCell.ImageLayout = DataGridViewImageCellLayout.Stretch;
                        iCell.Value = img;
                        dataGridView2[n, dataGridView2.RowCount - (i + 1)] = iCell;
                    }
                }
                //dataGridView2.Rows[5].Cells[4].Style.BackColor = Color.Green;
                //dataGridView2.Rows[5].Cells[4].
                MarkShelfs(z);
            }
            else {
                numericUpDown1.Value = 0;
                numericUpDown2.Value = 0;
                numericUpDown3.Value = 0;
            }
        }

        public DataGridView Frm1_DataGridViewShelf
        {
            get { return dataGridView2; }
            set { dataGridView2 = value; }
        }

        public class BoxLocation
        {
            public int boxn { get; set; }
            public string grey_id { get; set; }
            public int X { get; set; }
            public int Y { get; set; }
            public int Z { get; set; }
        }

        //List<BoxLocation> ShelfStructList = new List<BoxLocation>();
        List<BoxLocation> GetBoxesLocations(int z)
        {
            List<BoxLocation> list = new List<BoxLocation>();
            //
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT * FROM Boxes WHERE location_z='" + Convert.ToString(z) + "'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            BoxLocation box = new BoxLocation();
                            box.boxn = Convert.ToInt32(reader.GetValue(1).ToString());
                            box.grey_id = reader.GetValue(3).ToString();
                            box.X = Convert.ToInt32(reader.GetValue(4).ToString());
                            box.Y = Convert.ToInt32(reader.GetValue(5).ToString());
                            box.Z = Convert.ToInt32(reader.GetValue(6).ToString());
                            if (box.grey_id != "---")
                            {
                                list.Add(box);
                            }
                        }
                    }
                }
                connection.Close();
            }
            return list;
        }

        public void MarkShelfs(int z)
        {
            List<BoxLocation> box_locations = new List<BoxLocation>();
            //
            box_locations = GetBoxesLocations(z);
            //
            for (int i = 0; i < box_locations.Count; i++) {
                int x = box_locations[i].X;
                int y = box_locations[i].Y;
                x = x - 1;
                y = dataGridView2.RowCount - y;
                // Mark not empty with color
                //dataGridView2.Rows[y].Cells[x].Style.BackColor = Color.Green;
                //
                // Mark not empty with image
                DataGridViewImageCell iCell = new DataGridViewImageCell();
                System.Drawing.Image img = Properties.Resources.shelf_full3;
                iCell.ImageLayout = DataGridViewImageCellLayout.Stretch;
                iCell.Value = img;
                dataGridView2[x, y] = iCell;
                //dataGridView2[1, 2] = iCell;
            }
        }

        // Оформление прихода товара на склад
        //
        private Excel.Application App;
        private Excel.Range rng = null;
        private bool finish = false;
        DataTable mainTable = new DataTable();
        public void constructDatatable()
        {
            mainTable.Columns.Add("Number of Box", typeof(string));
            mainTable.Columns.Add("Description of Good", typeof(string));
            mainTable.Columns.Add("Quantity", typeof(string));
            mainTable.Columns.Add("N-weight", typeof(string));
            mainTable.Columns.Add("G-weight", typeof(string));
            dataGridView2.DataSource = mainTable;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        public void SelectPackList()
        {
            // Загружает в datagridview packing list в разделе 'Оформить приход'
            //
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(openFileDialog.FileName);
                ExcelPackage excel = new ExcelPackage(file);
                var worksheet = excel.Workbook.Worksheets[1];

                dataGridView3.RowCount = worksheet.Dimension.Rows;
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView3.Rows[k].Cells[0].Value = worksheet.Cells[k + 2, 1].Value;
                    dataGridView3.Rows[k].Cells[1].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView3.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView3.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 4].Value;
                    dataGridView3.Rows[k].Cells[4].Value = worksheet.Cells[k + 2, 5].Value;
                    dataGridView3.Rows[k].Cells[5].Value = worksheet.Cells[k + 2, 6].Value;
                }
                //dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                int i = 0;
                for (i = 0; i < dataGridView3.Columns.Count; i++)
                {
                    if (dataGridView3.Columns[i].HeaderText.ToString().Length < 0)
                    {
                        dataGridView3.Columns.RemoveAt(i);
                    }
                }

                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    int k = dataGridView3.Rows.IndexOf(row);
                    if (k % 2 == 0)
                    {
                        row.DefaultCellStyle.BackColor = cl_even;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = cl_odd;
                    }
                    row.HeaderCell.Value = "";
                }
                //
                dataGridView3.RowHeadersWidth = 35;
                //
                dataGridView3.Columns[0].HeaderText = "# коробки";
                dataGridView3.Columns[1].HeaderText = "Название продукта";
                dataGridView3.Columns[2].HeaderText = "Код продукта";
                dataGridView3.Columns[3].HeaderText = "Кол-во";
                dataGridView3.Columns[4].HeaderText = "Вес NET";
                dataGridView3.Columns[5].HeaderText = "Вес GROSS";

                dataGridView3.Columns[0].Width = 85;
                dataGridView3.Columns[1].Width = 330;
                dataGridView3.Columns[2].Width = 100;
                dataGridView3.Columns[3].Width = 70;
                dataGridView3.Columns[4].Width = 100;
                dataGridView3.Columns[5].Width = 100;

                int indexOfTotalRow = 0;
                indexOfTotalRow = dataGridView3.RowCount - 1;
                int qty = 0;
                double netWeight = 0;
                double grossWeight = 0;
                for(int k=0; k<dataGridView3.RowCount - 1; k++)
                {
                    qty = qty + Convert.ToInt32(dataGridView3.Rows[k].Cells[3].Value);
                    netWeight = netWeight + Math.Round(Convert.ToDouble(dataGridView3.Rows[k].Cells[4].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    grossWeight = grossWeight + Math.Round(Convert.ToDouble(dataGridView3.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
            }
                
                dataGridView3.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
                dataGridView3.Rows[indexOfTotalRow].Cells[3].Value = qty;
                dataGridView3.Rows[indexOfTotalRow].Cells[4].Value = netWeight;
                dataGridView3.Rows[indexOfTotalRow].Cells[5].Value = grossWeight;
                dataGridView3.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;


                //
                button7_in_save.Enabled = true;
                button9_in_gendoc.Enabled = true;
            }
            openFileDialog.Dispose();
        }

        public void UploadPckingList()
        {
            int boxNo;
            string idOrder;
            string pl_id;
            string pl_date;
            string invoice_id;
            string invoice_date;
            string partName;
            string partCode;
            string amount;
            string netWeight;
            string grossWeight;
            string blackId = "";
            int boxNoKeep = 0;

            string locationOfBoxX = "---";
            string locationOfBoxY = "---";
            string locationOfBoxZ = "---";
            string greyId;
            //
            idOrder = textBox4.Text;
            pl_id = textBox7.Text;
            string year = Convert.ToString(dateTimePicker4.Value.Year);
            string month = Convert.ToString(dateTimePicker4.Value.Month);
            string day = Convert.ToString(dateTimePicker4.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            pl_date = day + "." + month + "." + year;
            //
            invoice_id = textBox9.Text;
            year = Convert.ToString(dateTimePicker5.Value.Year);
            month = Convert.ToString(dateTimePicker5.Value.Month);
            day = Convert.ToString(dateTimePicker5.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            invoice_date = day + "." + month + "." + year;
            //
            //string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            int i = 0;
            for (i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                boxNo = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);

                if (boxNo != boxNoKeep)
                {
                    int randomValue = 0;
                    bool checkRandomVariable = false;
                    while (checkRandomVariable == false)
                    {
                        Random rnd = new Random();
                        randomValue = rnd.Next(9999, 99999);
                        blackId = randomValue.ToString(); // generate unique id for box

                        SqlCommand check_Black_Id = new SqlCommand("SELECT COUNT(*) FROM [Boxes] WHERE ([black_id] = @blackid)", connection);
                        check_Black_Id.Parameters.AddWithValue("@blackid", blackId);
                        int idExist = (int)check_Black_Id.ExecuteScalar();

                        if (idExist > 0)
                        {
                            checkRandomVariable = false;
                        }
                        else
                        {
                            checkRandomVariable = true;
                        }
                    }

                    greyId = "---";
                    // save box on Boxes tble
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO Boxes(box_number, black_id, grey_id, location_x, location_y, location_z) VALUES(@boxnumber, @blackId, @greyId, @locationX, @locationY, @locationZ)";
                    command.Parameters.AddWithValue("@boxnumber", boxNo);
                    command.Parameters.AddWithValue("@blackId", blackId);
                    command.Parameters.AddWithValue("@greyId", greyId);
                    command.Parameters.AddWithValue("@locationX", locationOfBoxX);
                    command.Parameters.AddWithValue("@locationY", locationOfBoxY);
                    command.Parameters.AddWithValue("@locationZ", locationOfBoxZ);
                    command.ExecuteNonQuery();

                    partName = dataGridView3.Rows[i].Cells[1].Value.ToString();
                    partCode = dataGridView3.Rows[i].Cells[2].Value.ToString();
                    amount = dataGridView3.Rows[i].Cells[3].Value.ToString();
                    netWeight = dataGridView3.Rows[i].Cells[4].Value.ToString();
                    grossWeight = dataGridView3.Rows[i].Cells[5].Value.ToString();

                    SqlCommand command1 = new SqlCommand();
                    command1.Connection = connection;
                    command1.CommandType = CommandType.Text;
                    command1.CommandText = "INSERT INTO items_in_boxes(box_numb, part_name, part_code, amount, net, gross, id_order, black_id, pl_id, pl_date, invoice_id, invoice_date) VALUES(@boxnumb, @partname, @partcode, @amount, @net, @gross, @idorder, @blackid, @pl_id, @pl_date, @invoice_id, @invoice_date)";
                    command1.Parameters.AddWithValue("@boxnumb", boxNo);
                    command1.Parameters.AddWithValue("@partname", partName);
                    command1.Parameters.AddWithValue("@partcode", partCode);
                    command1.Parameters.AddWithValue("@amount", amount);
                    command1.Parameters.AddWithValue("@net", netWeight);
                    command1.Parameters.AddWithValue("@gross", grossWeight);
                    command1.Parameters.AddWithValue("@idorder", idOrder);
                    command1.Parameters.AddWithValue("@blackid", blackId);
                    command1.Parameters.AddWithValue("@pl_id", pl_id);
                    command1.Parameters.AddWithValue("@pl_date", pl_date);
                    command1.Parameters.AddWithValue("@invoice_id", invoice_id);
                    command1.Parameters.AddWithValue("@invoice_date", invoice_date);
                    command1.ExecuteNonQuery();
                    boxNoKeep = boxNo;
                }
                else
                {
                    partName = dataGridView3.Rows[i].Cells[1].Value.ToString();
                    partCode = dataGridView3.Rows[i].Cells[2].Value.ToString();
                    amount = dataGridView3.Rows[i].Cells[3].Value.ToString();
                    netWeight = dataGridView3.Rows[i].Cells[4].Value.ToString();
                    grossWeight = dataGridView3.Rows[i].Cells[5].Value.ToString();

                    SqlCommand command1 = new SqlCommand();
                    command1.Connection = connection;
                    command1.CommandType = CommandType.Text;
                    command1.CommandText = "INSERT INTO items_in_boxes(box_numb, part_name, part_code, amount, net, gross, id_order, black_id, pl_id, pl_date, invoice_id, invoice_date) VALUES(@boxnumb, @partname, @partcode, @amount, @net, @gross, @idorder, @blackid, @pl_id, @pl_date, @invoice_id, @invoice_date)";
                    command1.Parameters.AddWithValue("@boxnumb", boxNoKeep);
                    command1.Parameters.AddWithValue("@partname", partName);
                    command1.Parameters.AddWithValue("@partcode", partCode);
                    command1.Parameters.AddWithValue("@amount", amount);
                    command1.Parameters.AddWithValue("@net", netWeight);
                    command1.Parameters.AddWithValue("@gross", grossWeight);
                    command1.Parameters.AddWithValue("@idorder", idOrder);
                    command1.Parameters.AddWithValue("@blackid", blackId);
                    command1.Parameters.AddWithValue("@pl_id", pl_id);
                    command1.Parameters.AddWithValue("@pl_date", pl_date);
                    command1.Parameters.AddWithValue("@invoice_id", invoice_id);
                    command1.Parameters.AddWithValue("@invoice_date", invoice_date);
                    command1.ExecuteNonQuery();
                }

            }
            connection.Close();
            MessageBox.Show("Packing list успешно загружен в систему.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            button8_in_cancel.Enabled = true;
            button9_in_gendoc.Enabled = true;
            button6_in_select.Enabled = false;
            button6_in_add.Enabled = false;
            button7_in_save.Enabled = false;
        }
        ///
        /// <summary>
        /// 
        /// </summary>

        private void создатьЗаказToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowPanels("Заказ");
            textBox1.Clear();
            textBox2.Clear();
            Size ns = new Size();
            ns.Width = panel7.Width - 30;
            ns.Height = panel7.Height - ((panel7.Height / 2) - (panel7.Height / 11));
            //richTextBox1.Size = ns;
            Point np = new Point();
            np.Y = richTextBox1.Location.Y + richTextBox1.Size.Height + 15;
            np.X = 15;
            //button3.Location = np;
            checkBox1.Checked = false;
            checkBox2.Checked = true;
            // Init gridview
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 8;
            //
            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "Код продукта";
            dataGridView1.Columns[2].HeaderText = "Название продукта";
            dataGridView1.Columns[3].HeaderText = "Количество";
            dataGridView1.Columns[4].HeaderText = "";
            dataGridView1.Columns[5].HeaderText = "Не хватает";
            dataGridView1.Columns[6].HeaderText = "На складе";
            dataGridView1.Columns[7].HeaderText = "Заметки";
            //
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[4].Width = 40;
            dataGridView1.Columns[5].Width = 90;
            dataGridView1.Columns[6].Width = 90;
            dataGridView1.Columns[7].Width = (dataGridView1.Size.Width - (dataGridView1.Columns[0].Width + dataGridView1.Columns[1].Width + dataGridView1.Columns[2].Width + dataGridView1.Columns[3].Width + dataGridView1.Columns[4].Width + dataGridView1.Columns[5].Width + dataGridView1.Columns[6].Width)) - 100;
            // // //
            dataGridView4.RowCount = 0;
            dataGridView4.ColumnCount = 10;
            //
            dataGridView4.Columns[0].HeaderText = "#";
            dataGridView4.Columns[1].HeaderText = "PO от";
            dataGridView4.Columns[2].HeaderText = "PO ID";
            dataGridView4.Columns[3].HeaderText = "PO дата";
            dataGridView4.Columns[4].HeaderText = "Order ID";
            dataGridView4.Columns[5].HeaderText = "Order дата";
            dataGridView4.Columns[6].HeaderText = "CP ID";
            dataGridView4.Columns[7].HeaderText = "CP дата";
            dataGridView4.Columns[8].HeaderText = "CT ID";
            dataGridView4.Columns[9].HeaderText = "CT дата";
            //
            //dataGridView4.Columns[0].Width = 90;
            //dataGridView4.Columns[1].Width = 80;
            //dataGridView4.Columns[2].Width = 80;
            //dataGridView4.Columns[3].Width = 90;
            //dataGridView4.Columns[4].Width = 80;
            //dataGridView4.Columns[5].Width = 80;
            //dataGridView4.Columns[6].Width = 130;
            // // //
            dataGridView5.RowCount = 0;
            dataGridView5.ColumnCount = 2;
            //
            dataGridView5.Columns[0].HeaderText = "Версия";
            dataGridView5.Columns[1].HeaderText = "Дата";
            //
            dataGridView5.Columns[0].Width = 70;
            dataGridView5.Columns[1].Width = 70;
            // // //
            dataGridView6.RowCount = 0;
            dataGridView6.ColumnCount = 8;
            //
            dataGridView6.Columns[0].HeaderText = "#";
            dataGridView6.Columns[1].HeaderText = "Код продукта";
            dataGridView6.Columns[2].HeaderText = "Название продукта";
            dataGridView6.Columns[3].HeaderText = "Кол-во";
            dataGridView6.Columns[4].HeaderText = "";
            dataGridView6.Columns[5].HeaderText = "Не хватает";
            dataGridView6.Columns[6].HeaderText = "На складе";
            dataGridView6.Columns[7].HeaderText = "Заметки";
            //
            dataGridView6.Columns[0].Width = 50;
            dataGridView6.Columns[1].Width = 130;
            dataGridView6.Columns[2].Width = 200;
            dataGridView6.Columns[3].Width = 60;
            dataGridView6.Columns[4].Width = 30;
            dataGridView6.Columns[5].Width = 60;
            dataGridView6.Columns[6].Width = 60;
            int sz = dataGridView6.Columns[0].Width + dataGridView6.Columns[1].Width + dataGridView6.Columns[2].Width + dataGridView6.Columns[3].Width + dataGridView6.Columns[4].Width + dataGridView6.Columns[5].Width + dataGridView6.Columns[6].Width;
            dataGridView6.Columns[7].Width = dataGridView6.Size.Width - sz - 60;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //button2.Enabled = true; moved to SelectOrder function
            SelectOrder();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //button2.Enabled = false;
            //button1.Enabled = false;
            //button3.Enabled = true;
            //button14.Enabled = true;
            //button13.Enabled = false;
            //textBox1.ReadOnly = true;
            //textBox5.ReadOnly = true;
            //textBox6.ReadOnly = true;
            //dateTimePicker1.Enabled = false;    all moved to UploadOrder function
            if (dataGridView1.RowCount - 1 == 0) {
                MessageBox.Show("Заказ пустой.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else {
                UploadOrder();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            textBox1.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            dateTimePicker1.Enabled = true;
            textBox1.Clear();
            textBox2.Clear();
            textBox5.Clear();
            textBox6.Clear();
            richTextBox1.Clear();
            dataGridView1.RowCount = 0;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = false;
            button14.Enabled = true;
            button13.Enabled = false;
            button27.Enabled = true;
            button28.Enabled = true;
            comboBox5.Enabled = true;
            menuStrip1.Items[0].Enabled = false;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            label104.Text = "1";
            textBox1.Clear();
            textBox2.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox13.Clear();
            richTextBox1.Clear();
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.RowCount = 0;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button14.Enabled = false;
            button13.Enabled = true;
            button27.Enabled = false;
            button28.Enabled = false;
            comboBox5.Enabled = false;
            menuStrip1.Items[0].Enabled = true;
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox1.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox13.Enabled = false;
            dateTimePicker1.Enabled = false;
            dateTimePicker7.Enabled = false;
            dateTimePicker8.Enabled = false;
            checkBox2.Checked = false;
        }

        private void checkBox2_Click(object sender, EventArgs e)
        {
            textBox1.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            textBox13.Enabled = true;
            dateTimePicker1.Enabled = true;
            dateTimePicker7.Enabled = true;
            dateTimePicker8.Enabled = true;
            checkBox1.Checked = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GenerateOrder();
        }

        private void управлениеСтелажамиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowPanels("Стелажи");
            MakeShelfsStructure(0);
            panel16.Width = 15;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            numericUpDown1.Enabled = true;
            numericUpDown2.Enabled = true;
            numericUpDown3.Enabled = true;
            button5.Enabled = true;
            shelfs_list_builded = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //MakeShelfsStructure();
            int z = Convert.ToInt32(numericUpDown1.Value);
            int x = Convert.ToInt32(numericUpDown2.Value);
            int y = Convert.ToInt32(numericUpDown3.Value);
            UpdateShelfsStruct(z, x, y);
            MakeShelfsStructure(0);
            numericUpDown1.Enabled = false;
            numericUpDown2.Enabled = false;
            numericUpDown3.Enabled = false;
            button5.Enabled = false;
        }

        private void checkedListBox1_Click(object sender, EventArgs e)
        {
             // стелаж
            _current_shelf_Z = checkedListBox1.SelectedIndex + 1;
            label11.Text = Convert.ToString(_current_shelf_Z);
            for (int i = 0; i < checkedListBox1.Items.Count; i++) {
                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }
            int a = checkedListBox1.SelectedIndex;
            checkedListBox1.SetItemCheckState(a, CheckState.Checked);
            try {
                if (shelf_mode == "search")
                {
                    Frm2.TextBox_Z = Convert.ToString(_current_shelf_Z);
                    MakeShelfsStructure(_current_shelf_Z);
                }
                if (shelf_mode == "box_location")
                {
                    Frm2.TextBox_Z_nloc = Convert.ToString(_current_shelf_Z);
                }
            }
            catch { }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // ярус
            _current_shelf_Y = dataGridView2.RowCount - dataGridView2.CurrentCell.RowIndex;
            label9.Text = Convert.ToString(_current_shelf_Y);
            // секция
            _current_shelf_X = dataGridView2.CurrentCell.ColumnIndex + 1;
            label10.Text = Convert.ToString(_current_shelf_X);
            //
            try {
                string _csX = Convert.ToString(_current_shelf_X);
                string _csY = Convert.ToString(_current_shelf_Y);
                string _csZ = Convert.ToString(_current_shelf_Z);
                //
                // на форме поиска form2 переключает между режимами поиска, указания места на складе
                if (shelf_mode == "search") {
                    Frm2.TextBox_Y = Convert.ToString(_current_shelf_Y);
                    Frm2.TextBox_X = Convert.ToString(_current_shelf_X);
                    //
                    Frm2.constructDataGridByCoordinates2(_csX, _csY, _csZ);
                    string grey_id = Frm2.Frm2_DataGridView1.Rows[0].Cells[2].Value.ToString();
                    Frm2.selectBoxItems(grey_id);
                }
                if (shelf_mode == "box_location") {
                    Frm2.TextBox_Y_nloc = Convert.ToString(_current_shelf_Y);
                    Frm2.TextBox_X_nloc = Convert.ToString(_current_shelf_X);
                    //Frm2.selectBoxItems(grey_id);
                }
            }
            catch { }
        }

        private void оформитьПриходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowPanels("Приход");
            dataGridView3.RowCount = 0;
            dataGridView3.ColumnCount = 6;
            //
            dataGridView3.Columns[0].HeaderText = "# Коробки";
            dataGridView3.Columns[1].HeaderText = "Название продукта";
            dataGridView3.Columns[2].HeaderText = "Код продукта";
            dataGridView3.Columns[3].HeaderText = "Количество";
            dataGridView3.Columns[4].HeaderText = "Вес NET";
            dataGridView3.Columns[5].HeaderText = "Вес GROSS";

            dataGridView3.Columns[0].Width = 100;
            dataGridView3.Columns[1].Width = 330;
            dataGridView3.Columns[2].Width = 150;
            dataGridView3.Columns[3].Width = 100;
            dataGridView3.Columns[4].Width = 150;
            dataGridView3.Columns[5].Width = 150;
            //
            dataGridView7.ColumnCount = 6;
            dataGridView7.RowCount = 0;
            //
            dataGridView7.Columns[0].HeaderText = "PO ID";
            dataGridView7.Columns[1].HeaderText = "Order ID";
            dataGridView7.Columns[2].HeaderText = "ID pack. list";
            dataGridView7.Columns[3].HeaderText = "Дата pack. list";
            dataGridView7.Columns[4].HeaderText = "ID Invoice";
            dataGridView7.Columns[5].HeaderText = "Дата invoice";

            dataGridView7.Columns[0].Width = 80;
            dataGridView7.Columns[1].Width = 80;
            dataGridView7.Columns[2].Width = 90;
            dataGridView7.Columns[3].Width = 110;
            dataGridView7.Columns[4].Width = 90;
            dataGridView7.Columns[5].Width = 100;
            //
            dataGridView8.RowCount = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SelectPackList();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string year = Convert.ToString(dateTimePicker2.Value.Year);
            string month = Convert.ToString(dateTimePicker2.Value.Month);
            string day = Convert.ToString(dateTimePicker2.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string date = day + "." + month + "." + year;
            //
            if (comboBox1.Text == "PO дата" || comboBox1.Text == "Order дата" || comboBox1.Text == "CP дата" || comboBox1.Text == "CT дата") {
                GetOrdersList(date, comboBox1.Text);
            }
            else {
                GetOrdersList(textBox3.Text, comboBox1.Text);
            }
        }

        private void label19_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView5.RowCount = 0;
            dataGridView5.RowHeadersWidth = 35;
            dataGridView6.RowCount = 0;
            int selected = 0;
            try
            {
                selected = dataGridView4.CurrentCell.RowIndex;
                string id_order = dataGridView4.Rows[selected].Cells[2].Value.ToString();
                GetVersionsList(id_order);
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            int selected = 0;
            try     
            {
                selected = dataGridView4.CurrentCell.RowIndex;
                string id_order = dataGridView4.Rows[selected].Cells[2].Value.ToString();
                selected = dataGridView5.CurrentCell.RowIndex;
                string version = dataGridView5.Rows[selected].Cells[0].Value.ToString();
                LoadOrderItems(id_order, Convert.ToInt32(version));
                //MessageBox.Show(id_order, "Сообщение",
                //MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (dataGridView4.RowCount != 0) {
                dataGridView6.RowCount = 0;
                richTextBox2.Clear();
                dataGridView4.Enabled = false;
                dataGridView5.Enabled = false;
                button8_search.Enabled = false;
                button9_add_v.Enabled = false;
                button10_select_source.Enabled = true;
                button11_save_v.Enabled = true;
                button15_cancel.Enabled = true;
                button12_gen_doc.Enabled = false;
                button16_edit_v.Enabled = false;
                dateTimePicker3.Enabled = true;
                button29.Enabled = true;
                button30.Enabled = true;
                menuStrip1.Items[0].Enabled = false;
            }
            else {
                MessageBox.Show("Выберите ID заказа к которому необходимо добавить новую версию.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        private void button10_Click(object sender, EventArgs e)
        {
            SelectOrder_NewVersion();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dataGridView6.RowCount - 1 == 0)
            {
                MessageBox.Show("Заказ пустой.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                button10_select_source.Enabled = false;
                button11_save_v.Enabled = false;
                UploadOrder_NewVersion();
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            dataGridView4.Enabled = true;
            dataGridView5.Enabled = true;
            dataGridView6.AllowUserToAddRows = false;
            dataGridView6.RowCount = 0;
            richTextBox2.Clear();
            button8_search.Enabled = true;
            button9_add_v.Enabled = true;
            button12_gen_doc.Enabled = true;
            button16_edit_v.Enabled = true;
            button10_select_source.Enabled = false;
            button11_save_v.Enabled = false;
            button15_cancel.Enabled = false;
            button29.Enabled = false;
            button30.Enabled = false;
            menuStrip1.Items[0].Enabled = true;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (dataGridView4.RowCount != 0) {
                int selected4 = dataGridView4.CurrentCell.RowIndex;
                int selected5 = dataGridView5.CurrentCell.RowIndex;
                //
                string po = dataGridView4.Rows[selected4].Cells[2].Value.ToString();
                string po_date = dataGridView4.Rows[selected4].Cells[3].Value.ToString();
                string order = dataGridView4.Rows[selected4].Cells[4].Value.ToString();
                string order_date = dataGridView5.Rows[selected5].Cells[1].Value.ToString();
                string cp = dataGridView4.Rows[selected4].Cells[6].Value.ToString();
                string cp_date = dataGridView4.Rows[selected4].Cells[7].Value.ToString();
                string ct = dataGridView4.Rows[selected4].Cells[8].Value.ToString();
                string ct_date = dataGridView4.Rows[selected4].Cells[9].Value.ToString();
                string v = dataGridView5.Rows[selected5].Cells[0].Value.ToString();
                string company = dataGridView4.Rows[selected4].Cells[1].Value.ToString();
                //
                GenerateOrderHist(po, po_date, order, order_date, cp, cp_date, ct, ct_date, v, company);
            }
            else
            {
                MessageBox.Show("Выберите версию заказа.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            button29.Enabled = true;
            button30.Enabled = true;
            button15_cancel.Enabled = true;
            button11_save_v.Enabled = true;
            button9_add_v.Enabled = false;
            button8_search.Enabled = false;
            button10_select_source.Enabled = false;
            button12_gen_doc.Enabled = false;
            button16_edit_v.Enabled = false;
            dataGridView4.Enabled = false;
            dataGridView5.Enabled = false;
            menuStrip1.Items[0].Enabled = false;
        }

        private void button7_in_save_Click(object sender, EventArgs e)
        {
            UploadPckingList();
        }

        private void button8_in_cancel_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
            textBox7.Clear();
            textBox9.Clear();
            dataGridView3.RowCount = 0;
            menuStrip1.Items[0].Enabled = true;
            button7_in_save.Enabled = false;
            button9_in_gendoc.Enabled = false;
            button8_in_cancel.Enabled = false;
            button6_in_select.Enabled = false;
            button6_in_add.Enabled = true;
        }

        private void button6_in_add_Click(object sender, EventArgs e)
        {
            menuStrip1.Items[0].Enabled = false;
            button8_in_cancel.Enabled = true;
            button6_in_select.Enabled = true;
            button6_in_add.Enabled = false;
        }

        private void складToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowPanels("Стелажи");
            MakeShelfsStructure(0);
            panel16.Width = 15;
            //checkedListBox1.SetItemCheckState(0, CheckState.Checked);
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            int selected = 0;
            try {
                selected = dataGridView7.CurrentCell.RowIndex;
                string id_order = dataGridView7.Rows[selected].Cells[0].Value.ToString();
                if (checkBox5.Checked == true) {
                    GeneratePL(id_order, "print");
                }
                else {
                    GeneratePL(id_order, "");
                }
            }
            catch {
                MessageBox.Show("Укажите ID заказа (PO)", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            string year = Convert.ToString(dateTimePicker6.Value.Year);
            string month = Convert.ToString(dateTimePicker6.Value.Month);
            string day = Convert.ToString(dateTimePicker6.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string date = day + "." + month + "." + year;
            //
            dataGridView8.RowCount = 0;
            if (comboBox2.Text == "Дата packing list" || comboBox2.Text == "Дата invoice")
            {
                GetPackingList(date, comboBox2.Text);
            }
            else
            {
                GetPackingList(textBox8.Text, comboBox2.Text);
            }
        }
        
        private void dataGridView7_SelectionChanged_1(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView7.CurrentCell.RowIndex;
                string id_order = dataGridView7.Rows[selected].Cells[0].Value.ToString();
                LoadPackListItems(id_order);
            }
            catch { }
        }

        private void label25_Click(object sender, EventArgs e)
        {
            textBox8.Clear();
        }

        private void checkBox3_Click(object sender, EventArgs e)
        {
            checkBox4.Checked = false;
        }

        private void checkBox4_Click(object sender, EventArgs e)
        {
            checkBox3.Checked = false;
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true) {
                GenerateGreyID(Convert.ToInt32(numericUpDown4.Value.ToString()), "gen");
            }
            else {
                GenerateGreyID(Convert.ToInt32(numericUpDown4.Value.ToString()), "print");
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            label_boxn.Text = "- - - - -";
            label_boxn.Text = "- - - - - - - - - - - -";
            //string res = CheckOrderExists("190120131425");
            //string res = CheckOrderExists("000000000000");
            string res = CheckOrderExists(textBox10.Text);
            if (res != "") {
                button8_attach.Enabled = false;
                button9_attach_done.Enabled = true;
                textBox11.Enabled = true;
                textBox12.Enabled = true;
                textBox11.Focus();
            }
            else {
                MessageBox.Show("Заказа (PO) с указанным ID не существует.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            timer_scan.Enabled = true;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            string str = textBox12.Text;
            if (str.Length == 5)
            {
                label_boxn.Text = str;
            }
            if (str.Length == 12)
            {
                label_grey_id.Text = str;
            }
        }

        private void button9_attach_done_Click(object sender, EventArgs e)
        {
            if (label_boxn.Text != "- - - - - -" && label_grey_id.Text != "- - - - - - - - - - - -") {
                Attach_Black_Grey_ID(label_boxn.Text, label_grey_id.Text);
            }
            //textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            label_boxn.Text = "- - - - -";
            label_grey_id.Text = "- - - - - - - - - - - -";
            button8_attach.Enabled = true;
            button9_attach_done.Enabled = false;

        }

        private void timer_scan_Tick(object sender, EventArgs e)
        {
            //MessageBox.Show("Test.", "Сообщение",
            //MessageBoxButtons.OK, MessageBoxIcon.Information);
            string str = textBox11.Text;
            if (str.Length == 5)
            {
                label_boxn.Text = str;
            }
            if (str.Length == 12)
            {
                label_grey_id.Text = str;
            }
            if (textBox11.Text != "") {
                textBox12.Focus();
            }
            timer_scan.Enabled = false;
        }

        private void button8_Click_2(object sender, EventArgs e)
        {
            textBox11.Clear();
            textBox12.Clear();
            label_boxn.Text = "- - - - -";
            label_grey_id.Text = "- - - - - - - - - - - -";
            textBox11.Focus();
        }

        private void label30_Click(object sender, EventArgs e)
        {
            textBox10.Clear();
        }

        private void checkBox5_Click(object sender, EventArgs e)
        {
            checkBox6.Checked = false;
        }

        private void checkBox6_Click(object sender, EventArgs e)
        {
            checkBox5.Checked = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Frm2.Show();
        }

        private void button10_Click_2(object sender, EventArgs e)
        {
            Frm3.Show();
        }

        private void label37_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void label36_Click(object sender, EventArgs e)
        {
            textBox6.Clear();
        }

        private void label34_Click(object sender, EventArgs e)
        {
            textBox5.Clear();
        }

        private void label35_Click(object sender, EventArgs e)
        {
            textBox13.Clear();
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            Frm4.Show();
        }

        private void созданиеКомерческогоПредложенияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowPanels("Ком");
            InitGridView12();
            InitGridView9();
            InitGrid10();
            GetCPList("", 0);
            FindCompanies();
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
        }

        private void dataGridView9_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int column = e.ColumnIndex;
            CPGridCellChange(row, column);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            LoadCPGrid("");
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            toolStripMenuItem1.Enabled = true;
            button15.Enabled = false;
            button17.Enabled = true;
            button18.Enabled = true;
            button19.Enabled = true;
            button48.Enabled = true;
            contextMenuStrip3.Enabled = true;
            contextMenuStrip3.Items[0].Enabled = true;
            contextMenuStrip3.Items[1].Enabled = true;
            //dataGridView9.RowCount = 1;
            //dataGridView9.AllowUserToAddRows = true;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            toolStripMenuItem1.Enabled = false;
            button15.Enabled = true;
            button16.Enabled = false;
            button17.Enabled = false;
            button18.Enabled = false;
            button19.Enabled = false;
            button48.Enabled = false;
            contextMenuStrip3.Enabled = false;
            dataGridView9.RowCount = 0;
            //dataGridView9.AllowUserToAddRows = false;
        }

        private void button12_Click_2(object sender, EventArgs e)
        {
            // %
            if (textBox18.Text != "")
            {
                try
                {
                    double c = Math.Round(Convert.ToDouble(textBox18.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    RefreshCPGrid(textBox18.Text, 0);
                }
                catch
                {
                    MessageBox.Show("Указанное значение не является числом.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            // RW
            if (textBox19.Text != "")
            {
                try
                {
                    double c = Math.Round(Convert.ToDouble(textBox19.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    RefreshCPGrid(textBox19.Text, 1);
                }
                catch
                {
                    MessageBox.Show("Указанное значение не является числом.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            // AIR
            if (textBox20.Text != "")
            {
                try
                {
                    double c = Math.Round(Convert.ToDouble(textBox20.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    RefreshCPGrid(textBox20.Text, 2);
                }
                catch
                {
                    MessageBox.Show("Указанное значение не является числом.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            // Uncheck grid checkboxes
            checkAllCheckBox.Checked = false;
            //
            foreach (DataGridViewRow item in dataGridView9.Rows)
            {
                DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)item.Cells[0];
                //
                cell.Value = false;
            }
        }

        private void label38_Click(object sender, EventArgs e)
        {
            textBox18.Clear();
        }

        private void label39_Click(object sender, EventArgs e)
        {
            textBox16.Clear();
        }

        private void label44_Click(object sender, EventArgs e)
        {
            textBox17.Clear();
        }

        private void label46_Click(object sender, EventArgs e)
        {
            textBox15.Clear();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text != "") {
                if (dataGridView9.RowCount != 0) {
                    UploadCP();
                    GenerateCP(dataGridView9, comboBox3, 1);
                }
                else {
                    MessageBox.Show("Комерческое предложение не может быть пустое.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            } else {
                MessageBox.Show("Необходимо указать способ доставки.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void label48_Click(object sender, EventArgs e)
        {
            textBox14.Clear();
        }

        private void label53_Click(object sender, EventArgs e)
        {
            textBox19.Clear();
        }

        private void label54_Click(object sender, EventArgs e)
        {
            textBox20.Clear();
        }

        private void textBox18_Click(object sender, EventArgs e)
        {
            //textBox19.Clear();
            //textBox20.Clear();
        }

        private void textBox19_Click(object sender, EventArgs e)
        {
            //textBox18.Clear();
            //textBox20.Clear();
        }

        private void textBox20_Click(object sender, EventArgs e)
        {
            //textBox18.Clear();
            //textBox19.Clear();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex == 0 || comboBox4.SelectedIndex == -1)
            {
                GetCPList(textBox14.Text, 0);
            }
            if (comboBox4.SelectedIndex == 1)
            {
                string year = Convert.ToString(dateTimePicker9.Value.Year);
                string month = Convert.ToString(dateTimePicker9.Value.Month);
                string day = Convert.ToString(dateTimePicker9.Value.Day);
                if (month.Length == 1) { month = "0" + month; }
                if (day.Length == 1) { day = "0" + day; }
                string cp_date = day + "." + month + "." + year;
                GetCPList(cp_date, 1);
            }
        }

        public class ItemList
        {
            public string boxNumber { get; set; }
            public string partName { get; set; }
            public string partCode { get; set; }
            public string amount { get; set; }
            public string blackId { get; set; }
        }
        private void button27_Click(object sender, EventArgs e)
        {
            checkItems();
        }
        public void checkItems()
        {
            List<ItemList> items = new List<ItemList>();

            string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            string connectionString = "select box_numb, part_name, part_code, amount, black_id from items_in_boxes";
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = connectionString;
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
                });
            }
            string partCode = "";
            int quantity = 0;
            int checkQuantity = 0;
            int i = 0;
            for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                partCode = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                quantity = Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);

                int k = 0;
                for (k = 0; k < items.Count; k++)
                {
                    if (items[k].partCode == partCode)
                    {
                        checkQuantity = checkQuantity + Convert.ToInt32(items[k].amount);
                    }
                }

                if (checkQuantity >= quantity)
                {
                    dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Green;
                    dataGridView1.Rows[i].Cells[5].Value = "---";
                    dataGridView1.Rows[i].Cells[6].Value = checkQuantity;
                }
                else
                {
                    dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Red;
                    dataGridView1.Rows[i].Cells[5].Value = quantity - checkQuantity;
                    dataGridView1.Rows[i].Cells[6].Value = checkQuantity;
                }

                checkQuantity = 0;
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            clearTable();
        }

        public void clearTable()
        {
            dataGridView1.AllowUserToAddRows = false;
            int i = 0;
            for(i=0; i<dataGridView1.Rows.Count - 1; i++)
            {
                if(dataGridView1.Rows[i].Cells[5].Value.ToString() == "---")
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            int indexOfTotalRow = dataGridView1.RowCount - 1;
            int totalQty = 0;
            for(int k=0; k<indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[k].Cells[3].Value);
            }
            dataGridView1.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
           
        }

        private void button9_in_gendoc_Click(object sender, EventArgs e)
        {
            string id_order;
            id_order = textBox4.Text;
            GeneratePL(id_order, "");
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            comboBox5.Items.Clear();
            string cp_id = textBox5.Text;
            if(cp_id.Length==13)
            {
                SqlConnection connection = new SqlConnection(conString);
                connection.Open();
                SqlCommand comm = new SqlCommand();
                comm.Connection = connection;
                comm.CommandType = CommandType.Text;
                comm.CommandText = "SELECT version FROM com_proposal WHERE cp_id=@cpid ORDER BY version";
                comm.Parameters.AddWithValue("@cpid", cp_id);
                SqlDataReader reader = comm.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    comboBox5.Items.Insert(i, reader.GetValue(0).ToString());
                }
                connection.Close();
            }
            else
            {
                comboBox5.Items.Clear();
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cp_id = textBox5.Text;
            string version = comboBox5.GetItemText(comboBox5.SelectedItem);
            label104.Text = version;
            getItemsOfCp(cp_id, version);
        }

        public class cpItems
        {
            public string partCode { get; set; }
            public string partName { get; set; }
            public string quantity { get; set; }
        }

        public void getItemsOfCp(string cpId, string version)
        {
            List<cpItems> cp_items = new List<cpItems>();

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "SELECT part_code, part_name, quantity FROM items_in_cp WHERE cp_id=@cpid and version=@version";
            comm.Parameters.AddWithValue("@cpid", cpId);
            comm.Parameters.AddWithValue("@version", version);
            SqlDataReader reader = comm.ExecuteReader();
            int i = 0;
            while (reader.Read())
            {
                cp_items.Add(new cpItems()
                {
                    partCode = reader.GetValue(0).ToString(),
                    partName = reader.GetValue(1).ToString(),
                    quantity = reader.GetValue(2).ToString()
                });
            }
            connection.Close();

            setOrderByCp(cp_items);
        }

        public void setOrderByCp(List<cpItems> items)
        {
            dataGridView1.RowCount = items.Count+1;
            int i = 0;
            for(i=0; i<items.Count; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = i + 1;
                dataGridView1.Rows[i].Cells[1].Value = items[i].partCode;
                dataGridView1.Rows[i].Cells[2].Value = items[i].partName;
                dataGridView1.Rows[i].Cells[3].Value = items[i].quantity;
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            checkItemsOfVersion();
        }

        public void checkItemsOfVersion()
        {
            List<ItemList> items = new List<ItemList>();

            string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            string connectionString = "select box_numb, part_name, part_code, amount, black_id from items_in_boxes";
            SqlCommand comm = new SqlCommand();
            comm.Connection = connection;
            comm.CommandType = CommandType.Text;
            comm.CommandText = connectionString;
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
                });
            }
            string partCode = "";
            int quantity = 0;
            int checkQuantity = 0;
            int i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
            {
                partCode = Convert.ToString(dataGridView6.Rows[i].Cells[1].Value);
                quantity = Convert.ToInt32(dataGridView6.Rows[i].Cells[3].Value);

                int k = 0;
                for (k = 0; k < items.Count; k++)
                {
                    if (items[k].partCode == partCode)
                    {
                        checkQuantity = checkQuantity + Convert.ToInt32(items[k].amount);
                    }
                }

                if (checkQuantity >= quantity)
                {
                    dataGridView6.Rows[i].Cells[4].Style.BackColor = Color.Green;
                    dataGridView6.Rows[i].Cells[5].Value = "---";
                    dataGridView6.Rows[i].Cells[6].Value = checkQuantity;
                }
                else
                {
                    dataGridView6.Rows[i].Cells[4].Style.BackColor = Color.Red;
                    dataGridView6.Rows[i].Cells[5].Value = quantity - checkQuantity;
                    dataGridView6.Rows[i].Cells[6].Value = checkQuantity;
                }

                checkQuantity = 0;
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            clearTableOfVersion();
        }

        public void clearTableOfVersion()
        {
            dataGridView6.AllowUserToAddRows = false;
            int i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
            {
                if (dataGridView6.Rows[i].Cells[5].Value.ToString() == "---")
                {
                    dataGridView6.Rows.RemoveAt(i);
                    i--;
                }
            }
            int indexOfTotalRow = dataGridView6.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[k].Cells[3].Value);
            }
            dataGridView6.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
        }

        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                dataGridView11.RowCount = 0;
                dataGridView12.RowCount = 0;
                selected = dataGridView10.CurrentCell.RowIndex;
                string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
                GetCPVersions(cp_id);
            }
            catch { }
        }

        private void dataGridView11_SelectionChanged(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                selected = dataGridView10.CurrentCell.RowIndex;
                string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
                selected = dataGridView11.CurrentCell.RowIndex;
                string v = dataGridView11.Rows[selected].Cells[0].Value.ToString();
                string order_id = dataGridView11.Rows[selected].Cells[3].Value.ToString();
                string date = dataGridView11.Rows[selected].Cells[4].Value.ToString();
                string dt_d = date.Substring(0, 2);
                string dt_m = date.Substring(3, 2);
                string dt_y = date.Substring(6, 4);
                dateTimePicker16.Value = new DateTime(Convert.ToInt32(dt_y), Convert.ToInt32(dt_m), Convert.ToInt32(dt_d));
                textBox32.Text = order_id;
                GetCPItems(cp_id, Convert.ToInt32(v));
            }
            catch { }
        }

        private void dataGridView12_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int column = e.ColumnIndex;
            CPGridCellChange2(row, column);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            // %
            if (textBox25.Text != "")
            {
                try
                {
                    double c = Math.Round(Convert.ToDouble(textBox25.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    RefreshCPGrid2(textBox25.Text, 0);
                }
                catch
                {
                    MessageBox.Show("Указанное значение не является числом.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            // RW
            if (textBox24.Text != "")
            {
                try
                {
                    double c = Math.Round(Convert.ToDouble(textBox24.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    RefreshCPGrid2(textBox24.Text, 1);
                }
                catch
                {
                    MessageBox.Show("Указанное значение не является числом.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            // AIR
            if (textBox23.Text != "")
            {
                try
                {
                    double c = Math.Round(Convert.ToDouble(textBox23.Text.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    RefreshCPGrid2(textBox23.Text, 2);
                }
                catch
                {
                    MessageBox.Show("Указанное значение не является числом.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            //
            checkAllCheckBox2.Checked = false;
            //
            foreach (DataGridViewRow item in dataGridView12.Rows)
            {
                DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)item.Cells[0];
                //
                cell.Value = false;
            }
        }

        private void textBox25_Click(object sender, EventArgs e)
        {
            //textBox24.Clear();
            //textBox23.Clear();
        }

        private void textBox24_Click(object sender, EventArgs e)
        {
            //textBox25.Clear();
            //textBox23.Clear();
        }

        private void textBox23_Click(object sender, EventArgs e)
        {
            //textBox24.Clear();
            //textBox25.Clear();
        }

        private void label64_Click(object sender, EventArgs e)
        {
            textBox25.Clear();
        }

        private void label61_Click(object sender, EventArgs e)
        {
            textBox24.Clear();
        }

        private void label60_Click(object sender, EventArgs e)
        {
            textBox23.Clear();
        }

        private void label57_Click(object sender, EventArgs e)
        {
            textBox22.Clear();
        }

        private void label56_Click(object sender, EventArgs e)
        {
            textBox21.Clear();
        }

        private void button16_Click_1(object sender, EventArgs e)
        {
            GenerateCP(dataGridView9, comboBox3, 0);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            int selected = dataGridView11.CurrentCell.RowIndex;
            string d_t = dataGridView11.Rows[selected].Cells[2].Value.ToString();
            GenerateCP2(dataGridView12, d_t, 0);
        }

        private void button31_Click(object sender, EventArgs e)
        {
            contextMenuStrip2.Enabled = true;
            contextMenuStrip2.Items[0].Enabled = true;
            dataGridView10.Enabled = false;
            dataGridView11.Enabled = false;
            dataGridView12.RowCount = 0;
            button23.Enabled = true;
            button25.Enabled = true;
            button26.Enabled = true;
            button52.Enabled = true;
            button22.Enabled = false;
            button31.Enabled = false;
            groupBox14.Enabled = true;
            contextMenuStrip4.Enabled = true;
            contextMenuStrip4.Items[0].Enabled = true;
            contextMenuStrip4.Items[1].Enabled = true;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            contextMenuStrip2.Enabled = false;
            contextMenuStrip4.Enabled = false;
            dataGridView10.Enabled = true;
            dataGridView11.Enabled = true;
            dataGridView12.RowCount = 0;
            button23.Enabled = false;
            button25.Enabled = false;
            button26.Enabled = false;
            button52.Enabled = false;
            button22.Enabled = true;
            button31.Enabled = true;
            groupBox14.Enabled = false;
            textBox32.Clear();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            contextMenuStrip2.Enabled = true;
            contextMenuStrip2.Items[0].Enabled = true;
            dataGridView10.Enabled = false;
            dataGridView11.Enabled = false;
            button23.Enabled = true;
            button25.Enabled = true;
            button26.Enabled = true;
            button52.Enabled = true;
            button22.Enabled = false;
            button31.Enabled = false;
            groupBox14.Enabled = true;
            contextMenuStrip4.Enabled = true;
            contextMenuStrip4.Items[0].Enabled = true;
            contextMenuStrip4.Items[1].Enabled = true;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            LoadCPGrid2("");
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (comboBox6.Text != "")
            {
                if (dataGridView12.RowCount != 0)
                {
                    int selected = dataGridView10.CurrentCell.RowIndex;
                    string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
                    string customer = dataGridView10.Rows[selected].Cells[2].Value.ToString();
                    int v = dataGridView11.RowCount;
                    string prj = dataGridView10.Rows[selected].Cells[3].Value.ToString();
                    UploadCP2(cp_id, v, customer, prj);
                    GetCPVersions(cp_id);
                    // 
                    dataGridView11.CurrentCell = dataGridView11.Rows[dataGridView11.RowCount - 1].Cells[0];
                    dataGridView11.CurrentCell.Selected = true;
                    selected = 0;
                    try
                    {
                        selected = dataGridView11.CurrentCell.RowIndex;
                        string version = dataGridView11.Rows[selected].Cells[0].Value.ToString();
                        string order_id = dataGridView11.Rows[selected].Cells[3].Value.ToString();
                        string date = dataGridView11.Rows[selected].Cells[4].Value.ToString();
                        string dt_d = date.Substring(0, 2);
                        string dt_m = date.Substring(3, 2);
                        string dt_y = date.Substring(6, 4);
                        dateTimePicker16.Value = new DateTime(Convert.ToInt32(dt_y), Convert.ToInt32(dt_m), Convert.ToInt32(dt_d));
                        textBox32.Text = order_id;
                        GetCPItems(cp_id, Convert.ToInt32(version));
                    }
                    catch { }
                    //
                    selected = dataGridView11.CurrentCell.RowIndex;
                    string d_t = dataGridView11.Rows[selected].Cells[2].Value.ToString();
                    GenerateCP2(dataGridView12, d_t, 1);
                }
                else
                {
                    MessageBox.Show("Комерческое предложение не может быть пустое.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Необходимо указать способ доставки.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void label66_Click(object sender, EventArgs e)
        {
            textBox26.Clear();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            if (comboBox7.SelectedIndex == 0 || comboBox7.SelectedIndex == -1)
            {
                GetContractList(textBox28.Text, 0);
            }
            if (comboBox7.SelectedIndex == 1)
            {
                string year = Convert.ToString(dateTimePicker13.Value.Year);
                string month = Convert.ToString(dateTimePicker13.Value.Month);
                string day = Convert.ToString(dateTimePicker13.Value.Day);
                if (month.Length == 1) { month = "0" + month; }
                if (day.Length == 1) { day = "0" + day; }
                string contract_date = day + "." + month + "." + year;
                GetContractList(contract_date, 1);
            }
        }

        // axad 22.07 start
        public class Contracts
        {
            public string id_contract { get; set; }
            public string date { get; set; }
            public string delivery_type { get; set; }
            public string delivery_point { get; set; }
            public string customer { get; set; }
            public string amount { get; set; }
            public string version { get; set; }
        }

        public void GetContractList(string argument, int type)
        {
            // search argument cp_id
            if (type == 0)
            {
                if (argument == "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, customer, amount, id FROM contracts WHERE version=1 ORDER BY id", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.customer = reader.GetValue(1).ToString();
                                    item.amount = reader.GetValue(2).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView13.RowCount = 0;
                    dataGridView14.RowCount = 0;
                    dataGridView24.RowCount = 0;
                    dataGridView13.RowCount = contracts.Count + 1;
                    dataGridView13.RowHeadersWidth = 35;
                    double totalAmount = 0;

                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView13.Rows[i].Cells[0].Value = i + 1;
                        dataGridView13.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView13.Rows[i].Cells[2].Value = contracts[i].customer;
                        dataGridView13.Rows[i].Cells[3].Value = Math.Round(Convert.ToDouble(contracts[i].amount), 2, MidpointRounding.ToEven);
                        totalAmount = totalAmount + Math.Round(Convert.ToDouble(contracts[i].amount), 2, MidpointRounding.ToEven);
                    }

                    int indexOfTotalRow = dataGridView13.RowCount - 1;
                    dataGridView13.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL:";
                    dataGridView13.Rows[indexOfTotalRow].Cells[3].Value = totalAmount;
                    dataGridView13.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;

                    for (int k=0; k<dataGridView13.RowCount-1; k++)
                    {
                        if (k % 2 == 0)
                        {
                            dataGridView13.Rows[k].Cells[0].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[1].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[2].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView13.Rows[k].Cells[0].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[1].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[2].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
                //
                if (argument != "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, customer, amount, id FROM contracts WHERE version=1 AND id_contract ='" + argument + "'", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.customer = reader.GetValue(1).ToString();
                                    item.amount = reader.GetValue(2).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView13.RowCount = 0;
                    dataGridView14.RowCount = 0;
                    dataGridView24.RowCount = 0;
                    dataGridView13.RowCount = contracts.Count + 1;
                    dataGridView13.RowHeadersWidth = 35;
                    double totalAmount = 0;

                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView13.Rows[i].Cells[0].Value = i + 1;
                        dataGridView13.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView13.Rows[i].Cells[2].Value = contracts[i].customer;
                        dataGridView13.Rows[i].Cells[3].Value = Math.Round(Convert.ToDouble(contracts[i].amount), 2, MidpointRounding.ToEven);
                        totalAmount = totalAmount + Math.Round(Convert.ToDouble(contracts[i].amount), 2, MidpointRounding.ToEven);
                    }

                    int indexOfTotalRow = dataGridView13.RowCount - 1;
                    dataGridView13.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL:";
                    dataGridView13.Rows[indexOfTotalRow].Cells[3].Value = totalAmount;
                    dataGridView13.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;

                    for (int k = 0; k < dataGridView13.RowCount - 1; k++)
                    {
                        if (k % 2 == 0)
                        {
                            dataGridView13.Rows[k].Cells[0].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[1].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[2].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView13.Rows[k].Cells[0].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[1].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[2].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
            // search argument date
            if (type == 1)
            {
                //
                if (argument != "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, customer, amount, id FROM contracts WHERE version=1 AND date='" + argument + "' ORDER BY id", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.customer = reader.GetValue(1).ToString();
                                    item.amount = reader.GetValue(2).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView13.RowCount = 0;
                    dataGridView14.RowCount = 0;
                    dataGridView24.RowCount = 0;
                    dataGridView13.RowCount = contracts.Count + 1;
                    dataGridView13.RowHeadersWidth = 35;
                    double totalAmount = 0;

                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView13.Rows[i].Cells[0].Value = i + 1;
                        dataGridView13.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView13.Rows[i].Cells[2].Value = contracts[i].customer;
                        dataGridView13.Rows[i].Cells[3].Value = Math.Round(Convert.ToDouble(contracts[i].amount), 2, MidpointRounding.ToEven);
                        totalAmount = totalAmount + Math.Round(Convert.ToDouble(contracts[i].amount), 2, MidpointRounding.ToEven);
                    }

                    int indexOfTotalRow = dataGridView13.RowCount - 1;
                    dataGridView13.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL:";
                    dataGridView13.Rows[indexOfTotalRow].Cells[3].Value = totalAmount;
                    dataGridView13.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;

                    for (int k = 0; k < dataGridView13.RowCount - 1; k++)
                    {
                        if (k % 2 == 0)
                        {
                            dataGridView13.Rows[k].Cells[0].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[1].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[2].Style.BackColor = cl_even;
                            dataGridView13.Rows[k].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView13.Rows[k].Cells[0].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[1].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[2].Style.BackColor = cl_odd;
                            dataGridView13.Rows[k].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
        }
        // axad 22.07 end


        // axad 22.07 start
        private void dataGridView13_SelectionChanged(object sender, EventArgs e)
        {
            int selected = 0;
            try
            {
                dataGridView24.RowCount = 0;
                dataGridView14.RowCount = 0;
                selected = dataGridView13.CurrentCell.RowIndex;
                string id_contract = dataGridView13.Rows[selected].Cells[1].Value.ToString();
                GetContractVersions(id_contract);
            }
            catch { }
        }
        public void GetContractVersions(string contract_id)
        {
            List<Contracts> contract_versions = new List<Contracts>();

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT delivery_type, delivery_point_air, delivery_point_rw FROM contracts WHERE id_contract='" + contract_id + "' and version=" + 1, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            textBox41.Text = reader.GetValue(1).ToString();
                            textBox42.Text = reader.GetValue(2).ToString();
                        }
                    }
                }
                connection.Close();
            }

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT version, date, delivery_type, delivery_point_air, delivery_point_rw FROM contracts WHERE id_contract='" + contract_id + "' ORDER BY version", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Contracts item = new Contracts();
                            item.version = reader.GetValue(0).ToString();
                            item.date = reader.GetValue(1).ToString();
                            item.delivery_type = reader.GetValue(2).ToString();
                            if(reader.GetValue(2).ToString()=="AIR")
                            {
                                item.delivery_point = reader.GetValue(3).ToString();
                            }
                            if (reader.GetValue(2).ToString() == "RW")
                            {
                                item.delivery_point = reader.GetValue(4).ToString();
                            }
                            contract_versions.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            dataGridView24.RowCount = 0;
            dataGridView24.RowCount = contract_versions.Count;
            dataGridView24.RowHeadersWidth = 35;
            for (int i = 0; i < contract_versions.Count; i++)
            {
                dataGridView24.Rows[i].Cells[0].Value = contract_versions[i].version;
                dataGridView24.Rows[i].Cells[1].Value = contract_versions[i].delivery_type;
                dataGridView24.Rows[i].Cells[2].Value = contract_versions[i].delivery_point;
                dataGridView24.Rows[i].Cells[3].Value = contract_versions[i].date;

                if (i % 2 == 0)
                {
                    dataGridView24.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView24.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView24.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView24.Rows[i].Cells[3].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView24.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView24.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView24.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView24.Rows[i].Cells[3].Style.BackColor = cl_odd;
                }
            }
        }
        // axad 22.07 end

        public class ContractItems
        {
            public string id_contract { get; set; }
            public string item_code { get; set; }
            public string name { get; set; }
            public string quantity { get; set; }
            public string unit_price { get; set; }
            public string amount_price { get; set; }
            public string X { get; set; }
            public string Y { get; set; }
            public string Z { get; set; }
            public string B { get; set; }
        }

        //axad 22.07 start
        public void GetContractItems(string id_contract, string version)
        {
            List<ContractItems> contract_items = new List<ContractItems>();

            string pay_before_percent = "";
            string pay_before_period = "";
            string delivery_period = "";
            string terms_of_payment = "";
            string terms_of_delivery = "";
            int b = 0;
            using (SqlConnection connection1 = new SqlConnection(conString))
            {
                connection1.Open();
                using (SqlCommand command1 = new SqlCommand("SELECT pay_before_percent, pay_before_period, delivery_period, terms_of_payment, terms_of_delivery FROM contracts WHERE id_contract='"+id_contract+"' and version='"+version+"'", connection1))
                {
                    using (SqlDataReader reader1 = command1.ExecuteReader())
                    {
                        while (reader1.Read())
                        {
                            b++;
                            pay_before_percent = reader1.GetValue(0).ToString();
                            pay_before_period = reader1.GetValue(1).ToString();
                            delivery_period = reader1.GetValue(2).ToString();
                            terms_of_payment = reader1.GetValue(3).ToString();
                            terms_of_delivery = reader1.GetValue(4).ToString();
                        }
                    }
                }
                connection1.Close();
            }

            int a = b;
            textBox37.Text = pay_before_percent;
            textBox38.Text = pay_before_period;
            textBox39.Text = delivery_period;
            if (terms_of_payment == "1")
            {
                checkBox11.Checked = true;
            }
            if (terms_of_payment == "2")
            {
                checkBox12.Checked = true;
            }
            if (terms_of_payment == "3")
            {
                checkBox13.Checked = true;
            }
            if (terms_of_delivery == "1")
            {
                checkBox14.Checked = true;
            }
            if (terms_of_delivery == "2")
            {
                checkBox15.Checked = true;
            }

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT id_contract, item_code, name, quantity, unit_price, amount_price FROM items_in_contract WHERE id_contract='"+id_contract+"' and version='"+version+"'", connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ContractItems contract = new ContractItems();
                            contract.id_contract = reader.GetValue(0).ToString();
                            contract.item_code = reader.GetValue(1).ToString();
                            contract.name = reader.GetValue(2).ToString();
                            contract.quantity = reader.GetValue(3).ToString();
                            contract.unit_price = reader.GetValue(4).ToString();
                            contract.amount_price = reader.GetValue(5).ToString();
                            contract_items.Add(contract);
                        }
                    }
                }
                connection.Close();
            }

            dataGridView14.RowCount = 0;
            dataGridView14.RowCount = contract_items.Count + 1;
            dataGridView14.RowHeadersWidth = 35;
            for(int i = 0; i < contract_items.Count; i++)
            {
                dataGridView14.Rows[i].Cells[0].Value = i + 1;
                dataGridView14.Rows[i].Cells[1].Value = contract_items[i].item_code;
                dataGridView14.Rows[i].Cells[2].Value = contract_items[i].name;
                dataGridView14.Rows[i].Cells[3].Value = contract_items[i].quantity;
                dataGridView14.Rows[i].Cells[4].Value = Math.Round(Convert.ToDouble(contract_items[i].unit_price), 2, MidpointRounding.ToEven).ToString("0.00");
                dataGridView14.Rows[i].Cells[5].Value = Math.Round(Convert.ToDouble(contract_items[i].amount_price), 2, MidpointRounding.ToEven).ToString("0.00");
                if (i % 2 == 0)
                {
                    dataGridView14.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[5].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView14.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[5].Style.BackColor = cl_odd;
                }
            }
            int indexOfTotalRow = 0;
            indexOfTotalRow = dataGridView14.RowCount - 1;
            int qty = 0;
            double price = 0;
            double totalPrice = 0;
            for(int k=0; k<dataGridView14.RowCount-1; k++)
            {
                qty = qty + Convert.ToInt32(dataGridView14.Rows[k].Cells[3].Value.ToString());
                price = price + Math.Round(Convert.ToDouble(dataGridView14.Rows[k].Cells[4].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                totalPrice = totalPrice + Math.Round(Convert.ToDouble(dataGridView14.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
            }

            dataGridView14.Rows[indexOfTotalRow].Cells[2].Value = "TOTAL: ";
            dataGridView14.Rows[indexOfTotalRow].Cells[3].Value = qty;
            dataGridView14.Rows[indexOfTotalRow].Cells[4].Value = Math.Round(price, 2, MidpointRounding.ToEven).ToString("0.00");
            dataGridView14.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(totalPrice, 2, MidpointRounding.ToEven).ToString("0.00");
            dataGridView14.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
        }
        
        // axad 22.07 end

        private void button34_Click(object sender, EventArgs e)
        {
            if (dataGridView14.RowCount != 0) { 
                int selected = dataGridView13.CurrentCell.RowIndex;
                string id_contract = dataGridView13.Rows[selected].Cells[1].Value.ToString();
                string company = dataGridView13.Rows[selected].Cells[2].Value.ToString();

                int selected2 = dataGridView24.CurrentCell.RowIndex;
                string delivery_point = dataGridView24.Rows[selected2].Cells[2].Value.ToString();
                string delivery_type = dataGridView24.Rows[selected2].Cells[1].Value.ToString();
                string data = dataGridView24.Rows[selected2].Cells[3].Value.ToString();
                string version_contract = dataGridView24.Rows[selected2].Cells[0].Value.ToString();
                generateContract(id_contract, data, delivery_type, delivery_point, company, version_contract);
            }
        }

        public void generateContract(string id_contract, string data, string delivery_type, string delivery_point, string company, string version)
        {
            Document document = new Document();
            document.LoadFromFile(app_dir_temp+ "Contract_Sample.docx");

            document.Replace("id_contract_fromProject", id_contract, false, true);
            document.Replace("date_frP", data+".", false, true);
            document.Replace("company_name_fromProject", company, false, true);
            document.Replace("delivery_point_fromProject", delivery_point, false, true);
            document.Replace("pay_before_fromProject", textBox37.Text, false, true);
            document.Replace("period_fromProject", textBox38.Text, false, true);
            document.Replace("delivery_time_fromProject", textBox39.Text, false, true);
            document.Replace("period_inWords", NumberToWordsEnglish(Convert.ToInt32(textBox38.Text.ToString())), false, true);

            if(checkBox11.Checked==true)
            {
                document.Replace("option1_rus_fromProject", "перед отгрузкой", false, true);
                document.Replace("option1_eng_fromProject", "from the date of signing the contract", false, true);
            }
            if (checkBox12.Checked == true)
            {
                document.Replace("option1_rus_fromProject", "после подписания контракта", false, true);
                document.Replace("option1_eng_fromProject", "after signing the contract", false, true);
            }
            if (checkBox13.Checked == true)
            {
                document.Replace("option1_rus_fromProject", "после поставки в пункт назначения", false, true);
                document.Replace("option1_eng_fromProject", "after delivery to the destination", false, true);
            }

            if(checkBox14.Checked==true)
            {
                document.Replace("option2_rus1_fromProject", "после подписания контракта", false, true);
                document.Replace("pay_before_fromProject2", "", false, true);
                document.Replace("option2_rus2_fromProject", "", false, true);

                document.Replace("option2_eng1_fromProject", "after signing the contract", false, true);
                document.Replace("pay_before_fromProject2", "", false, true);
                document.Replace("option2_eng2_fromProject", "", false, true);

            }
            if (checkBox15.Checked == true)
            {
                document.Replace("option2_rus1_fromProject", "со дня осуществления", false, true);
                document.Replace("pay_before_fromProject2", textBox37.Text+" %", false, true);
                document.Replace("option2_rus2_fromProject", "предоплаты", false, true);

                document.Replace("option2_eng1_fromProject", "from the date of", false, true);
                document.Replace("pay_before_fromProject2", textBox37.Text+" %", false, true);
                document.Replace("option2_eng2_fromProject", "prepayment", false, true);
            }

            Section section = document.AddSection();



            Paragraph paragraph = section.AddParagraph();
            TextRange text = paragraph.AppendText("             Приложение №1 от " + data + " г. к контракту №"+id_contract+" от " + data + "г.");
            text.CharacterFormat.Bold = true;

            Paragraph paragraph3 = section.AddParagraph();
            TextRange text1 = paragraph3.AppendText("          Appendix №1 dd " + data + " to the contract №" + id_contract + " dd " + data);
            text1.CharacterFormat.Bold = true;

            Paragraph paragraph1 = section.AddParagraph();

            Paragraph paragraph2 = section.AddParagraph();
            TextRange text2 = paragraph2.AppendText("                                               Спецификация/Specification №1");
            text2.CharacterFormat.Bold = true;

            Paragraph paragraph4 = section.AddParagraph();

            double total = 0;
            int quantityTotal = 0;
            string header_priceOfOneItem = "";
            string header_amountPrice = "";
            if (delivery_type == "AIR")
            {
                header_priceOfOneItem = "Цена за шт. " + delivery_point + " Долл.США /\n Price per unit " + delivery_point + " USD";
                header_amountPrice = "Сумма " + delivery_point + " Долл.США /\n Amount " + delivery_point + " USD ";
            }
            if (delivery_type == "RW")
            {
                header_priceOfOneItem = "Цена за шт. " + delivery_point + " Долл.США /\n Price per unit CPT Tashkent, Uzbekistan, Sergeli stantion";
                header_amountPrice = "Сумма " + delivery_point + "  Долл.США /\n Amount " + delivery_point + " USD ";
            }
            String[] header = { "№", "Код продукта/\nItem code", "Наименование продукции/\nName and model", "Код ТНВЭД/\nHS Code", "Кол-во. шт/\nQt-y. рс", header_priceOfOneItem, header_amountPrice };

            Spire.Doc.Table table = section.AddTable();
            table.ResetCells(dataGridView14.RowCount + 1, 7);
            TableRow row = table.Rows[0];
            row.IsHeader = true;
            row.Height = 130;

            row.HeightType = TableRowHeightType.Exactly;

            for (int i = 0; i < header.Length; i++)
            {

                row.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                Paragraph p = row.Cells[i].AddParagraph();

                p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                TextRange txtRange = p.AppendText(header[i]);
                txtRange.CharacterFormat.Bold = true;
            }

            for (int k = 0; k < dataGridView14.RowCount - 1; k++)
            {
                string itemCode = "";
                string name = "";
                string hsCode = "";
                string quantity = "";
                string priceOfOneItem = "";
                string amountPrice = "";

                // axad 22.07 start
                TableRow dataRow = table.Rows[k + 1];
                dataRow.Height = 25;
                dataRow.HeightType = TableRowHeightType.Exactly;
                dataRow.RowFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Triple;

                dataRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                dataRow.Cells[0].AddParagraph().AppendText(dataGridView14.Rows[k].Cells[0].Value.ToString());
                dataRow.Cells[0].SetCellWidth(5, CellWidthType.Percentage);

                dataRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { itemCode = dataGridView14.Rows[k].Cells[1].Value.ToString(); } catch { itemCode = ""; }
                dataRow.Cells[1].AddParagraph().AppendText(itemCode);
                dataRow.Cells[1].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                dataRow.Cells[1].SetCellWidth(15, CellWidthType.Percentage);

                dataRow.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { name = dataGridView14.Rows[k].Cells[2].Value.ToString(); } catch { name = ""; }
                dataRow.Cells[2].AddParagraph().AppendText(name);
                dataRow.Cells[2].SetCellWidth(40, CellWidthType.Percentage);

                dataRow.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { hsCode = ""; } catch { hsCode = ""; }
                dataRow.Cells[3].AddParagraph().AppendText(hsCode);
                dataRow.Cells[3].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                dataRow.Cells[3].SetCellWidth(10, CellWidthType.Percentage);

                dataRow.Cells[4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { quantity = dataGridView14.Rows[k].Cells[3].Value.ToString(); } catch { quantity = ""; }
                dataRow.Cells[4].AddParagraph().AppendText(quantity);
                dataRow.Cells[4].SetCellWidth(10, CellWidthType.Percentage);
                dataRow.Cells[4].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                quantityTotal = quantityTotal + Convert.ToInt32(quantity);

                dataRow.Cells[5].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { priceOfOneItem = dataGridView14.Rows[k].Cells[4].Value.ToString(); } catch { priceOfOneItem = ""; }
                dataRow.Cells[5].AddParagraph().AppendText("$ " + priceOfOneItem);
                dataRow.Cells[5].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                dataRow.Cells[5].SetCellWidth(10, CellWidthType.Percentage);

                dataRow.Cells[6].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { amountPrice = dataGridView14.Rows[k].Cells[5].Value.ToString(); } catch { amountPrice = ""; }
                dataRow.Cells[6].AddParagraph().AppendText("$ " + amountPrice);
                dataRow.Cells[6].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                dataRow.Cells[6].SetCellWidth(10, CellWidthType.Percentage);

                total = total + Math.Round(Convert.ToDouble(amountPrice.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);

                // axad 22.07 end
            }

            TableRow rowTotal = table.Rows[table.Rows.Count - 1];
            total = Math.Round(total, 2);
            rowTotal.Height = 25;
            rowTotal.HeightType = TableRowHeightType.Exactly;
            rowTotal.RowFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Triple;

            rowTotal.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[0].AddParagraph().AppendText("");
            rowTotal.Cells[0].SetCellWidth(5, CellWidthType.Percentage);

            rowTotal.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[1].AddParagraph().AppendText("");
            rowTotal.Cells[1].SetCellWidth(15, CellWidthType.Percentage);

            rowTotal.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[2].AddParagraph().AppendText("Total:");
            rowTotal.Cells[2].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            rowTotal.Cells[2].SetCellWidth(40, CellWidthType.Percentage);

            rowTotal.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[3].AddParagraph().AppendText("");
            rowTotal.Cells[3].SetCellWidth(10, CellWidthType.Percentage);

            rowTotal.Cells[4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[4].AddParagraph().AppendText(quantityTotal.ToString());
            rowTotal.Cells[4].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            rowTotal.Cells[4].SetCellWidth(10, CellWidthType.Percentage);

            rowTotal.Cells[5].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[5].AddParagraph().AppendText("");
            rowTotal.Cells[5].SetCellWidth(10, CellWidthType.Percentage);

            rowTotal.Cells[6].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[6].AddParagraph().AppendText("$ " + total.ToString());
            rowTotal.Cells[6].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            rowTotal.Cells[6].SetCellWidth(10, CellWidthType.Percentage);


            int number = (int)total;
            string str = total.ToString();
            int index = str.IndexOf(".");
            string left = string.Empty;
            if(index!=-1)
            {
                left = str.Substring(index + 1);
            }
            else
            {
                left = "";
            }

            Paragraph paragraph5 = section.AddParagraph();

            Paragraph paragraph6 = section.AddParagraph();
            TextRange text6 = paragraph6.AppendText("   Итого: " + "$ " + total.ToString() + " (" + NumberToWordsEnglish(number) + ", " + left + ") долларов США. Ставка НДС - 0 %");
            text6.CharacterFormat.Bold = true;

            Paragraph paragraph7 = section.AddParagraph();
            TextRange text7 = paragraph7.AppendText("   Total: " + "$ " + total.ToString() + " (" + NumberToWordsEnglish(number) + ", " + left + ") US dollars. VAT rate is 0%.");
            text7.CharacterFormat.Bold = true;

            Paragraph paragraph8 = section.AddParagraph();
            TextRange text8 = paragraph8.AppendText("Условия поставки: " + delivery_point + " согласно правилам ИНКОТЕРМС 2010.");
            text8.CharacterFormat.Bold = true;

            Paragraph paragraph9 = section.AddParagraph();
            TextRange text9 = paragraph9.AppendText("Terms of delivery: " + delivery_point + " under the rules of «INCOTERMS - 2010».");
            text9.CharacterFormat.Bold = true;

            Paragraph paragraph10 = section.AddParagraph();
            TextRange text10 = paragraph10.AppendText("Страна происхождения: Корея / Country of origin: Korea.");
            text10.CharacterFormat.Bold = true;

            Paragraph paragraph11 = section.AddParagraph();
            TextRange text11 = paragraph11.AppendText("Производитель: / Manufacturer: «LSIS» Корея.Продукция соответствует ГОСТами ТУ: / Products meet state standards and specifications: IEC 62052.11.");
            text11.CharacterFormat.Bold = true;

            string finalPrice = "$ " + total.ToString() + " (" + NumberToWordsEnglish(number) + ", " + left + ")";
            document.Replace("price_fromProject", finalPrice, false, true);

            ParagraphStyle style = new ParagraphStyle(document);
            style.Name = "CustomStyle";
            style.CharacterFormat.FontName = "Arial";
            style.CharacterFormat.FontSize = 8;
            document.Styles.Add(style);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                {
                    TableCell cell = table.Rows[i].Cells[j];
                    cell.CellFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    cell.CellFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    cell.CellFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    foreach (Paragraph para in cell.Paragraphs)
                    {
                        para.ApplyStyle(style.Name);
                    }
                }
            }

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "DOCX (*.docx)|*.docx";
            saveFileDialog.FilterIndex = 3;
            saveFileDialog.FileName = company + "_" + id_contract + "-" + version;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                document.SaveToFile(saveFileDialog.FileName);
            }
        }

        public void generateContractFromCP(string idContract, string data, string delivery_type, string delivery_point, string company)
        {
            Document document = new Document();
            document.LoadFromFile(app_dir_temp + "Contract_Sample.docx");

            document.Replace("id_contract_fromProject", idContract, false, true);
            document.Replace("date_frP", data + ".", false, true);
            document.Replace("company_name_fromProject", company, false, true);
            document.Replace("delivery_point_fromProject", delivery_point, false, true);
            document.Replace("pay_before_fromProject", textBox44.Text, false, true);
            document.Replace("period_fromProject", textBox43.Text, false, true);
            document.Replace("delivery_time_fromProject", textBox40.Text, false, true);
            document.Replace("period_inWords", NumberToWordsEnglish(Convert.ToInt32(textBox43.Text.ToString())), false, true);

            if (checkBox20.Checked == true)
            {
                document.Replace("option1_rus_fromProject", "перед отгрузкой", false, true);
                document.Replace("option1_eng_fromProject", "from the date of signing the contract", false, true);
            }
            if (checkBox19.Checked == true)
            {
                document.Replace("option1_rus_fromProject", "после подписания контракта", false, true);
                document.Replace("option1_eng_fromProject", "after signing the contract", false, true);
            }
            if (checkBox18.Checked == true)
            {
                document.Replace("option1_rus_fromProject", "после поставки в пункт назначения", false, true);
                document.Replace("option1_eng_fromProject", "after delivery to the destination", false, true);
            }

            if (checkBox17.Checked == true)
            {
                document.Replace("option2_rus1_fromProject", "после подписания контракта", false, true);
                document.Replace("pay_before_fromProject2", "", false, true);
                document.Replace("option2_rus2_fromProject", "", false, true);

                document.Replace("option2_eng1_fromProject", "after signing the contract", false, true);
                document.Replace("pay_before_fromProject2", "", false, true);
                document.Replace("option2_eng2_fromProject", "", false, true);

            }
            if (checkBox16.Checked == true)
            {
                document.Replace("option2_rus1_fromProject", "со дня осуществления", false, true);
                document.Replace("pay_before_fromProject2", textBox44.Text + " %", false, true);
                document.Replace("option2_rus2_fromProject", "предоплаты", false, true);

                document.Replace("option2_eng1_fromProject", "from the date of", false, true);
                document.Replace("pay_before_fromProject2", textBox44.Text + " %", false, true);
                document.Replace("option2_eng2_fromProject", "prepayment", false, true);
            }


            Section section = document.AddSection();

            Paragraph paragraph = section.AddParagraph();
            TextRange text = paragraph.AppendText("             Приложение №1 от " + data + " г. к контракту №" + idContract + " от " + data + "г.");
            text.CharacterFormat.Bold = true;

            Paragraph paragraph3 = section.AddParagraph();
            TextRange text1 = paragraph3.AppendText("          Appendix №1 dd " + data + " to the contract №" + idContract + " dd " + data);
            text1.CharacterFormat.Bold = true;

            Paragraph paragraph1 = section.AddParagraph();

            Paragraph paragraph2 = section.AddParagraph();
            TextRange text2 = paragraph2.AppendText("                                               Спецификация/Specification №1");
            text2.CharacterFormat.Bold = true;

            Paragraph paragraph4 = section.AddParagraph();

            double total = 0;
            int quantityTotal = 0;
            string header_priceOfOneItem = "";
            string header_amountPrice = "";
            if (delivery_type == "AIR")
            {
                header_priceOfOneItem = "Цена за шт. " + delivery_point + " Долл.США /\n Price per unit " + delivery_point + " USD";
                header_amountPrice = "Сумма " + delivery_point + " Долл.США /\n Amount " + delivery_point + " USD ";
            }
            if (delivery_type == "RW")
            {
                header_priceOfOneItem = "Цена за шт. " + delivery_point + " Долл.США /\n Price per unit CPT Tashkent, Uzbekistan, Sergeli stantion";
                header_amountPrice = "Сумма " + delivery_point + "  Долл.США /\n Amount " + delivery_point + " USD ";
            }
            String[] header = { "№", "Код продукта/\nItem code", "Наименование продукции/\nName and model", "Код ТНВЭД/\nHS Code", "Кол-во. шт/\nQt-y. рс", header_priceOfOneItem, header_amountPrice };

            Spire.Doc.Table table = section.AddTable();
            table.ResetCells(dataGridView12.RowCount + 1, 7);
            TableRow row = table.Rows[0];
            row.IsHeader = true;
            row.Height = 130;

            row.HeightType = TableRowHeightType.Exactly;

            for (int i = 0; i < header.Length; i++)
            {

                row.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                Paragraph p = row.Cells[i].AddParagraph();

                p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                TextRange txtRange = p.AppendText(header[i]);
                txtRange.CharacterFormat.Bold = true;
            }

            for (int k = 0; k < dataGridView12.RowCount - 1; k++)
            {
                string itemCode = "";
                string name = "";
                string hsCode = "";
                string quantity = "";
                string priceOfOneItem = "";
                string amountPrice = "";

                TableRow dataRow = table.Rows[k + 1];
                dataRow.Height = 25;
                dataRow.HeightType = TableRowHeightType.Exactly;
                dataRow.RowFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Triple;

                dataRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                dataRow.Cells[0].AddParagraph().AppendText(dataGridView12.Rows[k].Cells[1].Value.ToString());
                dataRow.Cells[0].SetCellWidth(5, CellWidthType.Percentage);

                dataRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { itemCode = dataGridView12.Rows[k].Cells[2].Value.ToString(); } catch { itemCode = ""; }
                dataRow.Cells[1].AddParagraph().AppendText(itemCode);
                dataRow.Cells[1].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                dataRow.Cells[1].SetCellWidth(15, CellWidthType.Percentage);

                dataRow.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { name = dataGridView12.Rows[k].Cells[3].Value.ToString(); } catch { name = ""; }
                dataRow.Cells[2].AddParagraph().AppendText(name);
                dataRow.Cells[2].SetCellWidth(40, CellWidthType.Percentage);

                dataRow.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { hsCode = ""; } catch { hsCode = ""; }
                dataRow.Cells[3].AddParagraph().AppendText(hsCode);
                dataRow.Cells[3].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                dataRow.Cells[3].SetCellWidth(15, CellWidthType.Percentage);

                dataRow.Cells[4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                try { quantity = dataGridView12.Rows[k].Cells[5].Value.ToString(); } catch { quantity = ""; }
                dataRow.Cells[4].AddParagraph().AppendText(quantity);
                dataRow.Cells[4].SetCellWidth(10, CellWidthType.Percentage);
                dataRow.Cells[4].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                quantityTotal = quantityTotal + Convert.ToInt32(quantity);

                if (delivery_type == "AIR")
                {
                    dataRow.Cells[5].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    try { priceOfOneItem = dataGridView12.Rows[k].Cells[20].Value.ToString(); } catch { priceOfOneItem = ""; }
                    dataRow.Cells[5].AddParagraph().AppendText("$ " + priceOfOneItem);
                    dataRow.Cells[5].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    dataRow.Cells[5].SetCellWidth(10, CellWidthType.Percentage);

                    dataRow.Cells[6].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    try { amountPrice = dataGridView12.Rows[k].Cells[22].Value.ToString(); } catch { amountPrice = ""; }
                    dataRow.Cells[6].AddParagraph().AppendText("$ " + amountPrice);
                    dataRow.Cells[6].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    dataRow.Cells[6].SetCellWidth(10, CellWidthType.Percentage);
                }
                if (delivery_type == "RW")
                {
                    dataRow.Cells[5].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    try { priceOfOneItem = dataGridView12.Rows[k].Cells[19].Value.ToString(); } catch { priceOfOneItem = ""; }
                    dataRow.Cells[5].AddParagraph().AppendText("$ " + priceOfOneItem);
                    dataRow.Cells[5].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    dataRow.Cells[5].SetCellWidth(10, CellWidthType.Percentage);

                    dataRow.Cells[6].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    try { amountPrice = dataGridView12.Rows[k].Cells[21].Value.ToString(); } catch { amountPrice = ""; }
                    dataRow.Cells[6].AddParagraph().AppendText("$ " + amountPrice);
                    dataRow.Cells[6].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    dataRow.Cells[6].SetCellWidth(10, CellWidthType.Percentage);
                }

                total = total + Math.Round(Convert.ToDouble(amountPrice.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
            }

            TableRow rowTotal = table.Rows[table.Rows.Count - 1];
            total = Math.Round(total, 2);
            rowTotal.Height = 25;
            rowTotal.HeightType = TableRowHeightType.Exactly;
            rowTotal.RowFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Triple;

            rowTotal.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[0].AddParagraph().AppendText("");
            rowTotal.Cells[0].SetCellWidth(5, CellWidthType.Percentage);

            rowTotal.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[1].AddParagraph().AppendText("");
            rowTotal.Cells[1].SetCellWidth(15, CellWidthType.Percentage);

            rowTotal.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[2].AddParagraph().AppendText("Total:");
            rowTotal.Cells[2].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            rowTotal.Cells[2].SetCellWidth(40, CellWidthType.Percentage);

            rowTotal.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[3].AddParagraph().AppendText("");
            rowTotal.Cells[3].SetCellWidth(15, CellWidthType.Percentage);

            rowTotal.Cells[4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[4].AddParagraph().AppendText(quantityTotal.ToString());
            rowTotal.Cells[4].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            rowTotal.Cells[4].SetCellWidth(10, CellWidthType.Percentage);

            rowTotal.Cells[5].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[5].AddParagraph().AppendText("");
            rowTotal.Cells[5].SetCellWidth(10, CellWidthType.Percentage);

            rowTotal.Cells[6].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            rowTotal.Cells[6].AddParagraph().AppendText("$ " + total.ToString());
            rowTotal.Cells[6].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            rowTotal.Cells[6].SetCellWidth(10, CellWidthType.Percentage);


            int number = (int)total;
            string str = total.ToString();
            int index = str.IndexOf(".");
            string left = string.Empty;
            if (index != -1)
            {
                left = str.Substring(index + 1);
            }
            else
            {
                left = "";
            }

            Paragraph paragraph5 = section.AddParagraph();

            Paragraph paragraph6 = section.AddParagraph();
            TextRange text6 = paragraph6.AppendText("   Итого: " + "$ " + total.ToString() + " (" + NumberToWordsEnglish(number) + ", " + left + ") долларов США. Ставка НДС - 0 %");
            text6.CharacterFormat.Bold = true;

            Paragraph paragraph7 = section.AddParagraph();
            TextRange text7 = paragraph7.AppendText("   Total: " + "$ " + total.ToString() + " (" + NumberToWordsEnglish(number) + ", " + left + ") US dollars. VAT rate is 0%.");
            text7.CharacterFormat.Bold = true;

            Paragraph paragraph8 = section.AddParagraph();
            TextRange text8 = paragraph8.AppendText("Условия поставки: " + delivery_point + " согласно правилам ИНКОТЕРМС 2010.");
            text8.CharacterFormat.Bold = true;

            Paragraph paragraph9 = section.AddParagraph();
            TextRange text9 = paragraph9.AppendText("Terms of delivery: " + delivery_point + " under the rules of «INCOTERMS - 2010».");
            text9.CharacterFormat.Bold = true;

            Paragraph paragraph10 = section.AddParagraph();
            TextRange text10 = paragraph10.AppendText("Страна происхождения: Корея / Country of origin: Korea.");
            text10.CharacterFormat.Bold = true;

            Paragraph paragraph11 = section.AddParagraph();
            TextRange text11 = paragraph11.AppendText("Производитель: / Manufacturer: «LSIS» Корея.Продукция соответствует ГОСТами ТУ: / Products meet state standards and specifications: IEC 62052.11.");
            text11.CharacterFormat.Bold = true;

            string finalPrice = "$ " + total.ToString() + " (" + NumberToWordsEnglish(number) + ", " + left + ")";
            document.Replace("price_fromProject", finalPrice, false, true);

            ParagraphStyle style = new ParagraphStyle(document);
            style.Name = "CustomStyle";
            style.CharacterFormat.FontName = "Arial";
            style.CharacterFormat.FontSize = 8;
            document.Styles.Add(style);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                {
                    TableCell cell = table.Rows[i].Cells[j];
                    cell.CellFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    cell.CellFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    cell.CellFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                    foreach (Paragraph para in cell.Paragraphs)
                    {
                        para.ApplyStyle(style.Name);
                    }
                }
            }

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "DOCX (*.docx)|*.docx";
            saveFileDialog.FilterIndex = 3;
            saveFileDialog.FileName = company + "_" + idContract;

            string folder = contextMenuStrip7.Items[1].Text.Substring(7);
            document.SaveToFile(folder + "\\" + company + "_" + idContract + "-1.docx");

            dataGridView11.RowCount = 0;
            int selected = 0;
            selected = dataGridView10.CurrentCell.RowIndex;
            string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
            GetCPVersions(cp_id);

            textBox40.Text = "";
            textBox43.Text = "";
            textBox44.Text = "";
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            checkBox19.Checked = false;
            checkBox20.Checked = false;

            System.Diagnostics.Process.Start(folder + "\\" + company + "_" + idContract + "-1.docx");
        }

        public static string NumberToWordsEnglish(int number)
        {
            if (number == 0)
                return "zero";

            if (number < 0)
                return "minus " + NumberToWordsEnglish(Math.Abs(number));

            string words = "";

            if ((number / 1000000) > 0)
            {
                words += NumberToWordsEnglish(number / 1000000) + " million ";
                number %= 1000000;
            }
            if ((number / 1000) > 0)
            {
                words += NumberToWordsEnglish(number / 1000) + " thousand ";
                number %= 1000;
            }
            if ((number / 100) > 0)
            {
                words += NumberToWordsEnglish(number / 100) + " hundred ";
                number %= 100;
            }
            if (number > 0)
            {
                if (words != "")
                    words += "and ";

                var unitsMap = new[] { "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen" };
                var tensMap = new[] { "zero", "ten", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

                if (number < 20)
                    words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0)
                        words += "-" + unitsMap[number % 10];
                }
            }
            return words;
        }
        public static string NumberToWordsRussian(int number)
        {
            if (number == 0)
                return "zero";

            if (number < 0)
                return "minus " + NumberToWordsRussian(Math.Abs(number));

            string words = "";

            if ((number / 1000000) > 0)
            {
                words += NumberToWordsRussian(number / 1000000) + " миллион ";
                number %= 1000000;
            }
            if ((number / 1000) > 0)
            {
                words += NumberToWordsRussian(number / 1000) + " тысяч ";
                number %= 1000;
            }
            if ((number / 100) > 0)
            {
                words += NumberToWordsRussian(number / 100) + " сот ";
                number %= 100;
            }
            if (number > 0)
            {
                if (words != "")
                    words += "and ";

                var unitsMap = new[] { "ноль", "один", "два", "три", "четыри", "пять", "шесть", "семь", "восемь", "девять", "десять", "одинадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать" };
                var tensMap = new[] { "ноль", "десять", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семдесят", "восемдесят", "девяносто" };

                if (number < 20)
                    words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0)
                        words += "-" + unitsMap[number % 10];
                }
            }
            return words;
        }

        private void сборЗаказовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowPanels("Корзина");
            if (dataGridView15.RowCount == 0) {
                InitGrid();
            }
        }

        public void LoadCTListToCollect()
        {
            try {
                int selected = dataGridView15.CurrentCell.RowIndex;
                string id_contract = dataGridView15.Rows[selected].Cells[1].Value.ToString();
                string data = dataGridView15.Rows[selected].Cells[2].Value.ToString();
                string customer = dataGridView15.Rows[selected].Cells[3].Value.ToString();
                bool check = true;
                string id_contract_check = "";
                for (int i = 0; i < dataGridView19.RowCount; i++)
                {
                    id_contract_check = dataGridView19.Rows[i].Cells[1].Value.ToString();
                    if (id_contract == id_contract_check)
                    {
                        check = false;
                    }
                }
                if (check == true)
                {
                    dataGridView19.RowCount = dataGridView19.RowCount + 1;
                    dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[0].Value = dataGridView19.RowCount;
                    dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[1].Value = id_contract;
                    dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[2].Value = data;
                    dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[3].Value = customer;
                    if (dataGridView19.RowCount % 2 != 0)
                    {
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[0].Style.BackColor = cl_even;
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[1].Style.BackColor = cl_even;
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[2].Style.BackColor = cl_even;
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[3].Style.BackColor = cl_even;
                    }
                    else
                    {
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[0].Style.BackColor = cl_odd;
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[1].Style.BackColor = cl_odd;
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[2].Style.BackColor = cl_odd;
                        dataGridView19.Rows[dataGridView19.RowCount - 1].Cells[3].Style.BackColor = cl_odd;
                    }
                }
            }
            catch { }
        }

        public void LoadCTItems()
        {
            List<ContractItems> ct_items = new List<ContractItems>();
            string sql = "SELECT id_contract, item_code, name, quantity FROM items_in_contract WHERE ";
            for (int k = 0; k < dataGridView19.RowCount; k++)
            {
                sql = sql + "id_contract='" + dataGridView19.Rows[k].Cells[1].Value + "' or ";
            }
            string sql_res = sql.Substring(0, sql.Length - 4);
            sql_res = sql_res + " ORDER BY id_contract";
            //
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql_res, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ContractItems item = new ContractItems();
                            item.id_contract = reader.GetValue(0).ToString();
                            item.item_code = reader.GetValue(1).ToString();
                            item.name = reader.GetValue(2).ToString();
                            item.quantity = reader.GetValue(3).ToString();
                            ct_items.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            dataGridView16.RowCount = 0;
            dataGridView16.RowCount = ct_items.Count;
            dataGridView16.RowHeadersWidth = 35;
            int n = 0;
            int inc = 1;
            for (int i = 0; i < ct_items.Count; i++)
            {
                try
                {
                    if (ct_items[i].id_contract != ct_items[i - 1].id_contract)
                    {
                        dataGridView16.Rows[n + 1].Cells[0].Style.BackColor = Color.DarkGray;
                        dataGridView16.Rows[n + 1].Cells[1].Style.BackColor = Color.DarkGray;
                        dataGridView16.Rows[n + 1].Cells[2].Style.BackColor = Color.DarkGray;
                        dataGridView16.Rows[n + 1].Cells[3].Style.BackColor = Color.DarkGray;
                        dataGridView16.Rows[n + 1].Cells[4].Style.BackColor = Color.DarkGray;
                        dataGridView16.Rows[n + 1].Cells[5].Style.BackColor = Color.DarkGray;
                        dataGridView16.Rows[n + 1].Cells[6].Style.BackColor = Color.DarkGray;
                        //
                        n = n + 3;
                        inc = 1;
                        dataGridView16.RowCount = dataGridView16.RowCount + 3;
                    }
                }
                catch { }
                //
                dataGridView16.Rows[n].Cells[0].Value = inc;
                dataGridView16.Rows[n].Cells[1].Value = ct_items[i].id_contract;
                dataGridView16.Rows[n].Cells[2].Value = ct_items[i].item_code;
                dataGridView16.Rows[n].Cells[3].Value = ct_items[i].name;
                dataGridView16.Rows[n].Cells[4].Value = ct_items[i].quantity;
                dataGridView16.Rows[n].Cells[5].Value = ct_items[i].quantity;
                if (n % 2 == 0)
                {
                    dataGridView16.Rows[n].Cells[0].Style.BackColor = cl_even;
                    dataGridView16.Rows[n].Cells[1].Style.BackColor = cl_even;
                    dataGridView16.Rows[n].Cells[2].Style.BackColor = cl_even;
                    dataGridView16.Rows[n].Cells[3].Style.BackColor = cl_even;
                    dataGridView16.Rows[n].Cells[4].Style.BackColor = cl_even;
                    dataGridView16.Rows[n].Cells[5].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView16.Rows[n].Cells[0].Style.BackColor = cl_odd;
                    dataGridView16.Rows[n].Cells[1].Style.BackColor = cl_odd;
                    dataGridView16.Rows[n].Cells[2].Style.BackColor = cl_odd;
                    dataGridView16.Rows[n].Cells[3].Style.BackColor = cl_odd;
                    dataGridView16.Rows[n].Cells[4].Style.BackColor = cl_odd;
                    dataGridView16.Rows[n].Cells[5].Style.BackColor = cl_odd;
                }
                dataGridView16.Rows[n].Cells[6].Style.BackColor = Color.Red;
                n++;
                inc++;
            }
            findBoxes();
        }

        private void button36_Click(object sender, EventArgs e)
        {
            if (comboBox8.SelectedIndex == 0)
            {
                GetContractListToCollect(textBox29.Text, 0);
            }
            if (comboBox8.SelectedIndex == 1)
            {
                string year = Convert.ToString(dateTimePicker14.Value.Year);
                string month = Convert.ToString(dateTimePicker14.Value.Month);
                string day = Convert.ToString(dateTimePicker14.Value.Day);
                if (month.Length == 1) { month = "0" + month; }
                if (day.Length == 1) { day = "0" + day; }
                string contract_date = day + "." + month + "." + year;
                GetContractListToCollect(contract_date, 1);
            }
        }

        public void GetContractListToCollect(string argument, int type)
        {
            // search argument cp_id
            if (type == 0)
            {
                if (argument == "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, customer, id FROM contracts ORDER BY id", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.date = reader.GetValue(1).ToString();
                                    item.customer = reader.GetValue(2).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView15.RowCount = 0;
                    dataGridView15.RowCount = contracts.Count;
                    dataGridView15.RowHeadersWidth = 35;
                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView15.Rows[i].Cells[0].Value = i + 1;
                        dataGridView15.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView15.Rows[i].Cells[2].Value = contracts[i].date;
                        dataGridView15.Rows[i].Cells[3].Value = contracts[i].customer;
                        if (i % 2 == 0)
                        {
                            dataGridView15.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[2].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView15.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[2].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
                //
                if (argument != "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, customer FROM contracts WHERE id_contract ='" + argument + "'", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.date = reader.GetValue(1).ToString();
                                    item.customer = reader.GetValue(2).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView15.RowCount = 0;
                    dataGridView15.RowCount = contracts.Count;
                    dataGridView15.RowHeadersWidth = 35;
                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView15.Rows[i].Cells[0].Value = i + 1;
                        dataGridView15.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView15.Rows[i].Cells[2].Value = contracts[i].date;
                        dataGridView15.Rows[i].Cells[3].Value = contracts[i].customer;

                        if (i % 2 == 0)
                        {
                            dataGridView15.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[2].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView15.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[2].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
            // search argument date
            if (type == 1)
            {
                //
                if (argument != "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, customer, id FROM contracts WHERE date='" + argument + "' ORDER BY id", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.date = reader.GetValue(1).ToString();
                                    item.customer = reader.GetValue(2).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView15.RowCount = 0;
                    dataGridView15.RowCount = contracts.Count;
                    dataGridView15.RowHeadersWidth = 35;
                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView15.Rows[i].Cells[0].Value = i + 1;
                        dataGridView15.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView15.Rows[i].Cells[2].Value = contracts[i].date;
                        dataGridView15.Rows[i].Cells[3].Value = contracts[i].customer;
                        if (i % 2 == 0)
                        {
                            dataGridView15.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[2].Style.BackColor = cl_even;
                            dataGridView15.Rows[i].Cells[3].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView15.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[2].Style.BackColor = cl_odd;
                            dataGridView15.Rows[i].Cells[3].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
        }

        public void InitGridPO()
        {
            textBox1.Clear();
            textBox2.Clear();
            Size ns = new Size();
            ns.Width = panel7.Width - 30;
            ns.Height = panel7.Height - ((panel7.Height / 2) - (panel7.Height / 11));
            //richTextBox1.Size = ns;
            Point np = new Point();
            np.Y = richTextBox1.Location.Y + richTextBox1.Size.Height + 15;
            np.X = 15;
            //button3.Location = np;
            checkBox1.Checked = false;
            checkBox2.Checked = true;
            // Init gridview
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 8;
            //
            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "Код продукта";
            dataGridView1.Columns[2].HeaderText = "Название продукта";
            dataGridView1.Columns[3].HeaderText = "Количество";
            dataGridView1.Columns[4].HeaderText = "";
            dataGridView1.Columns[5].HeaderText = "Не хватает";
            dataGridView1.Columns[6].HeaderText = "На складе";
            dataGridView1.Columns[7].HeaderText = "Заметки";
            //
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[4].Width = 40;
            dataGridView1.Columns[5].Width = 90;
            dataGridView1.Columns[6].Width = 90;
            dataGridView1.Columns[7].Width = (dataGridView1.Size.Width - (dataGridView1.Columns[0].Width + dataGridView1.Columns[1].Width + dataGridView1.Columns[2].Width + dataGridView1.Columns[3].Width + dataGridView1.Columns[4].Width + dataGridView1.Columns[5].Width + dataGridView1.Columns[6].Width)) - 100;
            // // //
            dataGridView4.RowCount = 0;
            dataGridView4.ColumnCount = 10;
            //
            dataGridView4.Columns[0].HeaderText = "#";
            dataGridView4.Columns[1].HeaderText = "PO от";
            dataGridView4.Columns[2].HeaderText = "PO ID";
            dataGridView4.Columns[3].HeaderText = "PO дата";
            dataGridView4.Columns[4].HeaderText = "Order ID";
            dataGridView4.Columns[5].HeaderText = "Order дата";
            dataGridView4.Columns[6].HeaderText = "CP ID";
            dataGridView4.Columns[7].HeaderText = "CP дата";
            dataGridView4.Columns[8].HeaderText = "CT ID";
            dataGridView4.Columns[9].HeaderText = "CT дата";
            //
            //dataGridView4.Columns[0].Width = 90;
            //dataGridView4.Columns[1].Width = 80;
            //dataGridView4.Columns[2].Width = 80;
            //dataGridView4.Columns[3].Width = 90;
            //dataGridView4.Columns[4].Width = 80;
            //dataGridView4.Columns[5].Width = 80;
            //dataGridView4.Columns[6].Width = 130;
            // // //
            dataGridView5.RowCount = 0;
            dataGridView5.ColumnCount = 2;
            //
            dataGridView5.Columns[0].HeaderText = "Версия";
            dataGridView5.Columns[1].HeaderText = "Дата";
            //
            dataGridView5.Columns[0].Width = 70;
            dataGridView5.Columns[1].Width = 70;
            // // //
            dataGridView6.RowCount = 0;
            dataGridView6.ColumnCount = 8;
            //
            dataGridView6.Columns[0].HeaderText = "#";
            dataGridView6.Columns[1].HeaderText = "Код продукта";
            dataGridView6.Columns[2].HeaderText = "Название продукта";
            dataGridView6.Columns[3].HeaderText = "Кол-во";
            dataGridView6.Columns[4].HeaderText = "";
            dataGridView6.Columns[5].HeaderText = "Не хватает";
            dataGridView6.Columns[6].HeaderText = "На складе";
            dataGridView6.Columns[7].HeaderText = "Заметки";
            //
            dataGridView6.Columns[0].Width = 50;
            dataGridView6.Columns[1].Width = 130;
            dataGridView6.Columns[2].Width = 200;
            dataGridView6.Columns[3].Width = 60;
            dataGridView6.Columns[4].Width = 30;
            dataGridView6.Columns[5].Width = 90;
            dataGridView6.Columns[6].Width = 90;
            int sz = dataGridView6.Columns[0].Width + dataGridView6.Columns[1].Width + dataGridView6.Columns[2].Width + dataGridView6.Columns[3].Width + dataGridView6.Columns[4].Width + dataGridView6.Columns[5].Width + dataGridView6.Columns[6].Width;
            dataGridView6.Columns[7].Width = dataGridView6.Size.Width - sz - 60;
        }

        public void InitGridIn()
        {
            dataGridView3.RowCount = 0;
            dataGridView3.ColumnCount = 6;
            //
            dataGridView3.Columns[0].HeaderText = "# Коробки";
            dataGridView3.Columns[1].HeaderText = "Название продукта";
            dataGridView3.Columns[2].HeaderText = "Код продукта";
            dataGridView3.Columns[3].HeaderText = "Количество";
            dataGridView3.Columns[4].HeaderText = "Вес NET";
            dataGridView3.Columns[5].HeaderText = "Вес GROSS";

            dataGridView3.Columns[0].Width = 100;
            dataGridView3.Columns[1].Width = 330;
            dataGridView3.Columns[2].Width = 150;
            dataGridView3.Columns[3].Width = 100;
            dataGridView3.Columns[4].Width = 150;
            dataGridView3.Columns[5].Width = 150;
            //
            dataGridView7.ColumnCount = 6;
            dataGridView7.RowCount = 0;
            //
            dataGridView7.Columns[0].HeaderText = "PO ID";
            dataGridView7.Columns[1].HeaderText = "Order ID";
            dataGridView7.Columns[2].HeaderText = "ID pack. list";
            dataGridView7.Columns[3].HeaderText = "Дата pack. list";
            dataGridView7.Columns[4].HeaderText = "ID Invoice";
            dataGridView7.Columns[5].HeaderText = "Дата invoice";

            dataGridView7.Columns[0].Width = 80;
            dataGridView7.Columns[1].Width = 80;
            dataGridView7.Columns[2].Width = 90;
            dataGridView7.Columns[3].Width = 110;
            dataGridView7.Columns[4].Width = 90;
            dataGridView7.Columns[5].Width = 100;
            //
            //
            dataGridView8.ColumnCount = 7;
            dataGridView8.RowCount = 0;
            //
            dataGridView8.Columns[0].HeaderText = "# Коробки";
            dataGridView8.Columns[1].HeaderText = "Код продукта";
            dataGridView8.Columns[2].HeaderText = "Название продукта";
            dataGridView8.Columns[3].HeaderText = "Кол-во";
            dataGridView8.Columns[4].HeaderText = "Вес NET";
            dataGridView8.Columns[5].HeaderText = "Вес GROSS";
            dataGridView8.Columns[6].HeaderText = "ID коробки";
            //
            dataGridView8.Columns[0].Width = 90;
            dataGridView8.Columns[1].Width = 100;
            dataGridView8.Columns[2].Width = 260;
            dataGridView8.Columns[3].Width = 80;
            dataGridView8.Columns[4].Width = 80;
            dataGridView8.Columns[5].Width = 90;
            dataGridView8.Columns[6].Width = 90;
        }

        public void InitGrid()
        {
            dataGridView15.RowCount = 0;
            dataGridView15.ColumnCount = 4;

            dataGridView15.Columns[0].HeaderText = "#";
            dataGridView15.Columns[1].HeaderText = "Id Контракта";
            dataGridView15.Columns[2].HeaderText = "Дата";
            dataGridView15.Columns[3].HeaderText = "Заказчик";

            dataGridView15.Columns[0].Width = 40;
            dataGridView15.Columns[1].Width = 120;
            dataGridView15.Columns[2].Width = 70;
            dataGridView15.Columns[3].Width = 120;

            foreach (DataGridViewColumn col in dataGridView15.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //
            dataGridView19.RowCount = 0;
            dataGridView19.ColumnCount = 4;

            dataGridView19.Columns[0].HeaderText = "#";
            dataGridView19.Columns[1].HeaderText = "Id Контракта";
            dataGridView19.Columns[2].HeaderText = "Дата";
            dataGridView19.Columns[3].HeaderText = "Заказчик";

            dataGridView19.Columns[0].Width = 40;
            dataGridView19.Columns[1].Width = 120;
            dataGridView19.Columns[2].Width = 70;
            dataGridView19.Columns[3].Width = 120;

            foreach (DataGridViewColumn col in dataGridView19.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            dataGridView17.RowCount = 0;
            dataGridView17.ColumnCount = 5;

            dataGridView17.Columns[0].HeaderText = "ID Коробки";
            dataGridView17.Columns[1].HeaderText = "Секция";
            dataGridView17.Columns[2].HeaderText = "Ярус";
            dataGridView17.Columns[3].HeaderText = "Стелаж";
            dataGridView17.Columns[4].HeaderText = "Блок";

            dataGridView17.Columns[0].Width = 90;
            dataGridView17.Columns[1].Width = 50;
            dataGridView17.Columns[2].Width = 50;
            dataGridView17.Columns[3].Width = 50;
            dataGridView17.Columns[4].Width = 50;

            foreach (DataGridViewColumn col in dataGridView17.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //
            dataGridView18.RowCount = 0;
            dataGridView18.ColumnCount = 4;

            dataGridView18.Columns[0].HeaderText = "Код продукта";
            dataGridView18.Columns[1].HeaderText = "Название продукта";
            dataGridView18.Columns[2].HeaderText = "Кол-во";
            dataGridView18.Columns[3].HeaderText = "Надо";

            dataGridView18.Columns[0].Width = 120;
            dataGridView18.Columns[1].Width = 230;
            dataGridView18.Columns[2].Width = 60;
            dataGridView18.Columns[3].Width = 60;

            foreach (DataGridViewColumn col in dataGridView18.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //
            dataGridView16.RowCount = 0;
            dataGridView16.ColumnCount = 11;

            dataGridView16.Columns[0].HeaderText = "#";
            dataGridView16.Columns[1].HeaderText = "ID Контракта";
            dataGridView16.Columns[2].HeaderText = "Код продукта";
            dataGridView16.Columns[3].HeaderText = "Название продукта";
            dataGridView16.Columns[4].HeaderText = "Кол-во";
            dataGridView16.Columns[5].HeaderText = "Осталось";
            dataGridView16.Columns[6].HeaderText = "";
            dataGridView16.Columns[7].HeaderText = "Стелаж";
            dataGridView16.Columns[8].HeaderText = "Ярус";
            dataGridView16.Columns[9].HeaderText = "Секция";
            dataGridView16.Columns[10].HeaderText = "Блок";

            dataGridView16.Columns[0].Width = 40;
            dataGridView16.Columns[1].Width = 120;
            dataGridView16.Columns[2].Width = 100;
            dataGridView16.Columns[3].Width = 150;
            dataGridView16.Columns[4].Width = 60;
            dataGridView16.Columns[5].Width = 70;
            dataGridView16.Columns[6].Width = 30;
            dataGridView16.Columns[7].Width = 30;
            dataGridView16.Columns[8].Width = 30;
            dataGridView16.Columns[9].Width = 30;
            dataGridView16.Columns[10].Width = 30;
            dataGridView16.Columns[10].Visible = false;
            dataGridView16.Columns[9].Visible = false;
            dataGridView16.Columns[8].Visible = false;
            dataGridView16.Columns[7].Visible = false;

            foreach (DataGridViewColumn col in dataGridView16.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //
            dataGridView20.RowCount = 0;
            dataGridView20.ColumnCount = 3;

            dataGridView20.Columns[0].HeaderText = "#";
            dataGridView20.Columns[1].HeaderText = "ID Контракта";
            dataGridView20.Columns[2].HeaderText = "Дата";

            dataGridView20.Columns[0].Width = 40;
            dataGridView20.Columns[1].Width = 120;
            dataGridView20.Columns[2].Width = 80;

            foreach (DataGridViewColumn col in dataGridView20.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //
            dataGridView21.RowCount = 0;
            dataGridView21.ColumnCount = 14;

            dataGridView21.Columns[0].HeaderText = "#";
            dataGridView21.Columns[1].HeaderText = "ID Контракта";
            dataGridView21.Columns[2].HeaderText = "ID заказа клиента";
            dataGridView21.Columns[3].HeaderText = "Дата заказа клиента";
            dataGridView21.Columns[4].HeaderText = "Код продукта";
            dataGridView21.Columns[5].HeaderText = "Название продукта";
            dataGridView21.Columns[6].HeaderText = "HS Код";
            dataGridView21.Columns[7].HeaderText = "Кол-во";
            dataGridView21.Columns[8].HeaderText = "Цена за ед.";
            dataGridView21.Columns[9].HeaderText = "Общая сумма";
            dataGridView21.Columns[10].HeaderText = "Net-Вес";
            dataGridView21.Columns[11].HeaderText = "Gross-Вес";
            dataGridView21.Columns[12].HeaderText = "ID Коробки";
            dataGridView21.Columns[13].HeaderText = "ID Палеты";

            dataGridView21.Columns[0].Width = 40;
            dataGridView21.Columns[1].Width = 120;
            dataGridView21.Columns[2].Width = 120;
            dataGridView21.Columns[3].Width = 140;
            dataGridView21.Columns[4].Width = 120;
            dataGridView21.Columns[5].Width = 225;
            dataGridView21.Columns[6].Width = 80;
            dataGridView21.Columns[7].Width = 80;
            dataGridView21.Columns[8].Width = 80;
            dataGridView21.Columns[9].Width = 95;
            dataGridView21.Columns[10].Width = 80;
            dataGridView21.Columns[11].Width = 80;
            dataGridView21.Columns[12].Width = 80;
            dataGridView21.Columns[13].Width = 80;

            foreach (DataGridViewColumn col in dataGridView21.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //
        }

        private void MergeCellsInRow(DataGridView Grid, int row, int col1, int col2)
        {
            Graphics g = Grid.CreateGraphics();
            Pen p = new Pen(Grid.GridColor);
            Rectangle r1 = Grid.GetCellDisplayRectangle(col1, row, true);
            Rectangle r2 = Grid.GetCellDisplayRectangle(col2, row, true);

            int recWidth = 0;
            string recValue = string.Empty;
            for (int i = col1; i <= col2; i++)
            {
                recWidth += Grid.GetCellDisplayRectangle(i, row, true).Width;
                if (Grid[i, row].Value != null)
                    recValue += Grid[i, row].Value.ToString() + " ";
            }
            Rectangle newCell = new Rectangle(r1.X, r1.Y, recWidth, r1.Height);
            g.FillRectangle(new SolidBrush(Grid.DefaultCellStyle.BackColor), newCell);
            g.DrawRectangle(p, newCell);
            g.DrawString(recValue, Grid.DefaultCellStyle.Font, new SolidBrush(Grid.DefaultCellStyle.ForeColor), newCell.X + 3, newCell.Y + 3);
        }

        private void MergeCellsInColumn(DataGridView Grid, int col, int row1, int row2)
        {
            Graphics g = dataGridView1.CreateGraphics();
            Pen p = new Pen(dataGridView1.GridColor);
            Rectangle r1 = dataGridView1.GetCellDisplayRectangle(col, row1, true);
            Rectangle r2 = dataGridView1.GetCellDisplayRectangle(col, row2, true);

            int recHeight = 0;
            string recValue = string.Empty;
            for (int i = row1; i <= row2; i++)
            {
                recHeight += dataGridView1.GetCellDisplayRectangle(col, i, true).Height;
                if (dataGridView1[col, i].Value != null)
                    recValue += dataGridView1[col, i].Value.ToString() + " ";
            }
            Rectangle newCell = new Rectangle(r1.X, r1.Y, r1.Width, recHeight);
            g.FillRectangle(new SolidBrush(dataGridView1.DefaultCellStyle.BackColor), newCell);
            g.DrawRectangle(p, newCell);
            g.DrawString(recValue, dataGridView1.DefaultCellStyle.Font, new SolidBrush(dataGridView1.DefaultCellStyle.ForeColor), newCell.X + 3, newCell.Y + 3);
        }

        public class ItemsInBoxes
        {
            public string partName { get; set; }
            public string partCode { get; set; }
            public string amount { get; set; }
            public string blackId { get; set; }
            public string locationX { get; set; }
            public string locationY { get; set; }
            public string locationZ { get; set; }
            public string block { get; set; }
        }
        public class Locations
        {
            public string boxNo { get; set; }
            public string locationX { get; set; }
            public string locationY { get; set; }
            public string locationZ { get; set; }
            public string block { get; set; }
        }
        public class BoxItems
        {
            public string itemCode { get; set; }
            public string name { get; set; }
            public string quantity { get; set; }
        }
        public class ItemsToDisplay
        {
            public string itemCode { get; set; }
            public string name { get; set; }
            public string quantity { get; set; }
            public string needed { get; set; }
        }


        public void findBoxes()
        {
            List<ItemsInBoxes> all_items = new List<ItemsInBoxes>();
            List<Locations> locations = new List<Locations>();

            string sql = "SELECT items_in_boxes.part_name, items_in_boxes.part_code, items_in_boxes.amount, items_in_boxes.black_id, Boxes.location_x, Boxes.location_y, Boxes.location_z, Boxes.block FROM Boxes join items_in_boxes on Boxes.grey_id = items_in_boxes.black_id WHERE items_in_boxes.amount>0";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ItemsInBoxes item = new ItemsInBoxes();
                            item.partName = reader.GetValue(0).ToString();
                            item.partCode = reader.GetValue(1).ToString();
                            item.amount = reader.GetValue(2).ToString();
                            item.blackId = reader.GetValue(3).ToString();
                            item.locationX = reader.GetValue(4).ToString();
                            item.locationY = reader.GetValue(5).ToString();
                            item.locationZ = reader.GetValue(6).ToString();
                            item.block = reader.GetValue(7).ToString();
                            all_items.Add(item);
                        }
                    }
                }
                connection.Close();
            }

            string itemCode = "";
            int need = 0;
            for (int i = 0; i < dataGridView16.RowCount; i++)
            {
                try { itemCode = dataGridView16.Rows[i].Cells[2].Value.ToString(); } catch { itemCode = ""; }
                try { need = Convert.ToInt32(dataGridView16.Rows[i].Cells[5].Value.ToString()); } catch { need = 0; }
                for (int k = 0; k < all_items.Count; k++)
                {
                    if (all_items[k].partCode == itemCode && need > 0)
                    {
                        Locations box_location = new Locations();
                        box_location.boxNo = all_items[k].blackId;
                        box_location.locationX = all_items[k].locationX;
                        box_location.locationY = all_items[k].locationY;
                        box_location.locationZ = all_items[k].locationZ;
                        box_location.block = all_items[k].block;
                        locations.Add(box_location);
                    }
                }
            }
            setLocations(locations);
        }

        //
        public void setLocations(List<Locations> location)
        {
            for (int k = 0; k < location.Count; k++)
            {
                string box = location[k].boxNo;
                for (int n = k + 1; n < location.Count; n++)
                {
                    if (location[k].boxNo == location[n].boxNo)
                    {
                        location.RemoveAt(n);
                        n--;
                    }
                }
            }
            dataGridView18.RowCount = 0;
            dataGridView17.RowCount = 0;
            dataGridView17.RowCount = location.Count;
            for (int i = 0; i < location.Count; i++)
            {
                dataGridView17.Rows[i].Cells[0].Value = location[i].boxNo;
                dataGridView17.Rows[i].Cells[1].Value = location[i].locationX;
                dataGridView17.Rows[i].Cells[2].Value = location[i].locationY;
                dataGridView17.Rows[i].Cells[3].Value = location[i].locationZ;
                dataGridView17.Rows[i].Cells[4].Value = location[i].block;
                if (i % 2 == 0)
                {
                    dataGridView17.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView17.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView17.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView17.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView17.Rows[i].Cells[4].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView17.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView17.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView17.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView17.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView17.Rows[i].Cells[4].Style.BackColor = cl_odd;
                }
            }
        }


        private void button37_Click(object sender, EventArgs e)
        {
            dataGridView16.RowCount = 0;
            dataGridView17.RowCount = 0;
            dataGridView18.RowCount = 0;
            dataGridView19.RowCount = 0;
        }

        private void button35_Click(object sender, EventArgs e)
        {
            LoadCTListToCollect();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            dataGridView16.RowCount = 0;
            dataGridView17.RowCount = 0;
            dataGridView18.RowCount = 0;
            LoadCTItems();
        }

        private void button39_Click(object sender, EventArgs e)
        {

        }


        private void dataGridView17_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView18.RowCount = 0;
            Frm5.dataGridView4.RowCount = 0;
            int selected = 0;
            try
            {
                selected = dataGridView17.CurrentCell.RowIndex;
                string boxNo = dataGridView17.Rows[selected].Cells[0].Value.ToString();
                getBoxItems(boxNo);
            }
            catch { }
            if (Frm5.dataGridView1.RowCount == 0)
            {
                Frm5.TextBox_Item_Code = "";
                Frm5.TextBox_Name = "";
                Frm5.quantity = 0;
                Frm5.box_number = "";
            }
        }

        public void getBoxItems(string boxNumber)
        {
            List<BoxItems> items = new List<BoxItems>();
            List<ItemsToDisplay> itemsDisplay = new List<ItemsToDisplay>();

            string sql = "SELECT part_name, part_code, amount FROM items_in_boxes WHERE black_id=" + boxNumber + " and amount>0";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            BoxItems item = new BoxItems();
                            item.itemCode = reader.GetValue(1).ToString();
                            item.name = reader.GetValue(0).ToString();
                            item.quantity = reader.GetValue(2).ToString();
                            items.Add(item);
                        }
                    }
                }
                connection.Close();
            }

            for (int i = 0; i < items.Count; i++)
            {
                string itemCode = items[i].itemCode;
                string itemCodeOnTable = "";
                int needed = 0;
                bool check = false;
                for(int k = 0; k<dataGridView16.Rows.Count; k++)
                {
                    try { itemCodeOnTable = dataGridView16.Rows[k].Cells[2].Value.ToString(); } catch { itemCodeOnTable = ""; }
                    if (itemCode==itemCodeOnTable)
                    {
                        needed = needed + Convert.ToInt32(dataGridView16.Rows[k].Cells[5].Value.ToString());
                        if(needed>0)
                        {
                            check = true;
                        }
                    }
                }
                if(check==true)
                {
                    ItemsToDisplay item = new ItemsToDisplay();
                    item.itemCode = items[i].itemCode;
                    item.name = items[i].name;
                    item.quantity = items[i].quantity;
                    item.needed = needed.ToString();
                    itemsDisplay.Add(item);
                }
            }
            displayBoxItems(itemsDisplay);
        }
        public void displayBoxItems(List<ItemsToDisplay> listItems)
        {
            dataGridView18.RowCount = listItems.Count;
            for(int i = 0; i<listItems.Count; i++)
            {
                dataGridView18.Rows[i].Cells[0].Value = listItems[i].itemCode;
                dataGridView18.Rows[i].Cells[1].Value = listItems[i].name;
                dataGridView18.Rows[i].Cells[2].Value = listItems[i].quantity;
                dataGridView18.Rows[i].Cells[3].Value = listItems[i].needed;
                if (i % 2 == 0)
                {
                    dataGridView18.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView18.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView18.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView18.Rows[i].Cells[3].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView18.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView18.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView18.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView18.Rows[i].Cells[3].Style.BackColor = cl_odd;
                }
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
        }

        public string MakeWhiteID()
        {
            // Генерация WHite ID
            string white_id;
            Random rnd = new Random();
            int valueFirst = rnd.Next(9999, 99999);
            white_id = valueFirst.ToString();
            return white_id;
        }

        public void GenerateWhiteID(int qty, string type, string ct_id, string customer)
        {
            // Генерация Grey ID
            Document document = new Document();
            //
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "DOCX (*.docx)|*.docx";
            saveFileDialog.FilterIndex = 3;
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            int i = 0;
            for (i = 0; i < qty; i++)
            {
                string w_id_ToPrint = "";
                bool check_w_id = false;
                while (check_w_id == false)
                {
                    w_id_ToPrint = MakeWhiteID();
                    SqlCommand check_white_Id = new SqlCommand("SELECT COUNT(*) FROM [collected_contract_items] WHERE ([white_id] = @wid and [id_contract] = @ctid)", connection);
                    check_white_Id.Parameters.AddWithValue("@wid", w_id_ToPrint);
                    check_white_Id.Parameters.AddWithValue("@ctid", ct_id);
                    int idExist = (int)check_white_Id.ExecuteScalar();
                    //
                    if (idExist > 0)
                    {
                        check_w_id = false;
                    }
                    else
                    {
                        check_w_id = true;
                    }
                }
                Section section = document.AddSection();
                // White ID
                Paragraph paragraph = section.AddParagraph();
                paragraph.AppendText("*" + w_id_ToPrint + "*");
                //
                ParagraphStyle style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "IDAutomationHC39M";
                style.CharacterFormat.FontSize = 50;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
                //
                // Space paragraph
                paragraph.Format.LineSpacing = 15;
                //
                // Contract ID bar code
                paragraph = section.AddParagraph();
                paragraph.AppendText("*" + ct_id + "*");
                //
                style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "IDAutomationHC39M";
                style.CharacterFormat.FontSize = 20;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
                //
                paragraph.Format.LineSpacing = 20;
                //
                // Contract ID bar code
                paragraph = section.AddParagraph();
                paragraph.AppendText(ct_id);
                //
                style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "Calibri (Body)";
                style.CharacterFormat.FontSize = 45;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
                //
                paragraph.Format.LineSpacing = 20;
                //
                // Customer
                paragraph = section.AddParagraph();
                paragraph.AppendText(customer);
                //
                style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "Calibri (Body)";
                style.CharacterFormat.FontSize = 45;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
            }
            //
            if (type == "print")
            {
                document.SaveToFile(app_dir_temp + "wh_id.docx");
                //
                PrintDialog printDlg = new PrintDialog();
                printDlg.AllowPrintToFile = true;
                //printDlg.AllowCurrentPage = true;
                //printDlg.AllowSelection = true;
                //printDlg.AllowSomePages = true;
                printDlg.UseEXDialog = true;
                //
                Document Doc = new Document();
                Doc.LoadFromFile(app_dir_temp + "wh_id.docx");
                Doc.PrintDialog = printDlg;
                //
                System.Drawing.Printing.PrintDocument printDoc = Doc.PrintDocument;
                //printDoc.Print();
                //
                if (printDlg.ShowDialog() == DialogResult.OK)
                {
                    printDoc.Print();
                }
                File.Delete(app_dir_temp + "grey_id.docx");
                //
            }
            else
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    document.SaveToFile(saveFileDialog.FileName);
                }
            }
            //
        }

        public string MakePaletID()
        {
            // Генерация WHite ID
            string white_id;
            Random rnd = new Random();
            int valueFirst = rnd.Next(999, 9999);
            white_id = valueFirst.ToString();
            return white_id;
        }

        public void GeneratePaletID(int qty, string type, string ct_id, string customer)
        {
            // Генерация Grey ID
            Document document = new Document();
            //
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "DOCX (*.docx)|*.docx";
            saveFileDialog.FilterIndex = 3;
            //
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();
            int i = 0;
            for (i = 0; i < qty; i++)
            {
                string pid = "";
                bool check_pid = false;
                while (check_pid == false){
                    pid = MakePaletID();
                    SqlCommand check_pid_com = new SqlCommand("SELECT COUNT(*) FROM [collected_contract_items] WHERE ([no_poleta] = @pid and [id_contract] = @ctid)", connection);
                    check_pid_com.Parameters.AddWithValue("@pid", pid);
                    check_pid_com.Parameters.AddWithValue("@ctid", ct_id);
                    int idExist = (int)check_pid_com.ExecuteScalar();
                    //
                    if (idExist > 0){
                        check_pid = false;
                    }else{
                        check_pid = true;
                    }
                }
                Section section = document.AddSection();
                // White ID
                Paragraph paragraph = section.AddParagraph();
                paragraph.AppendText("*" + pid + "*");
                //
                ParagraphStyle style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "IDAutomationHC39M";
                style.CharacterFormat.FontSize = 50;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
                //
                // Space paragraph
                paragraph.Format.LineSpacing = 15;
                //
                // Contract ID bar code
                //paragraph = section.AddParagraph();
                //paragraph.AppendText("*" + ct_id + "*");
                ////
                //style = new ParagraphStyle(document);
                //style.CharacterFormat.FontName = "IDAutomationHC39M";
                //style.CharacterFormat.FontSize = 20;
                //paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                ////
                //document.Styles.Add(style);
                //paragraph.ApplyStyle(style.Name);
                ////
                //paragraph.Format.LineSpacing = 20;
                //
                // Contract ID bar code
                paragraph = section.AddParagraph();
                paragraph.AppendText(ct_id);
                //
                style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "Calibri (Body)";
                style.CharacterFormat.FontSize = 45;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
                //
                paragraph.Format.LineSpacing = 20;
                //
                // Customer
                paragraph = section.AddParagraph();
                paragraph.AppendText(customer);
                //
                style = new ParagraphStyle(document);
                style.CharacterFormat.FontName = "Calibri (Body)";
                style.CharacterFormat.FontSize = 45;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //
                document.Styles.Add(style);
                paragraph.ApplyStyle(style.Name);
            }
            //
            if (type == "print")
            {
                document.SaveToFile(app_dir_temp + "pid.docx");
                //
                PrintDialog printDlg = new PrintDialog();
                printDlg.AllowPrintToFile = true;
                //printDlg.AllowCurrentPage = true;
                //printDlg.AllowSelection = true;
                //printDlg.AllowSomePages = true;
                printDlg.UseEXDialog = true;
                //
                Document Doc = new Document();
                Doc.LoadFromFile(app_dir_temp + "pid.docx");
                Doc.PrintDialog = printDlg;
                //
                System.Drawing.Printing.PrintDocument printDoc = Doc.PrintDocument;
                //printDoc.Print();
                //
                if (printDlg.ShowDialog() == DialogResult.OK)
                {
                    printDoc.Print();
                }
                File.Delete(app_dir_temp + "pid.docx");
                //
            }
            else
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    document.SaveToFile(saveFileDialog.FileName);
                }
            }
            //
        }

        private void dataGridView18_SelectionChanged(object sender, EventArgs e)
        {
            Frm5.dataGridView1.RowCount = 0;
            Frm5.dataGridView4.RowCount = 0;
            Frm5.dataGridView5.RowCount = 0;
            Frm5.Show();
            Frm5.Focus();
            int selected = 0;
            int selected_1 = 0;
            try
            {
                selected = dataGridView18.CurrentCell.RowIndex;
                string item_code = dataGridView18.Rows[selected].Cells[0].Value.ToString();
                string name = dataGridView18.Rows[selected].Cells[1].Value.ToString();
                int quantity_in_box =Convert.ToInt32(dataGridView18.Rows[selected].Cells[2].Value.ToString());
                selected_1 = dataGridView17.CurrentCell.RowIndex;
                string box_number = dataGridView17.Rows[selected_1].Cells[0].Value.ToString();

                Frm5.TextBox_Item_Code = item_code;
                Frm5.TextBox_Name = name;
                Frm5.quantity = quantity_in_box;
                Frm5.box_number = box_number;
                Frm5.comboBox1.Items.Clear();

                for(int i = 0; i < dataGridView16.RowCount; i++)
                {
                    string item_code_check = "";
                    try { item_code_check = dataGridView16.Rows[i].Cells[2].Value.ToString(); } catch { item_code_check = ""; }

                    string contract_id = "";
                    try { contract_id = dataGridView16.Rows[i].Cells[1].Value.ToString(); } catch { contract_id = ""; }

                    int needed = 0;
                    try { needed = Convert.ToInt32(dataGridView16.Rows[i].Cells[5].Value.ToString()); } catch { needed = 0; }

                    if (item_code_check == item_code && needed > 0)
                    {
                        //
                        Frm5.dataGridView1.RowCount = Frm5.dataGridView1.RowCount + 1;
                        Frm5.dataGridView1.RowHeadersWidth = 35;
                        //
                        Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[0].Value = Frm5.dataGridView1.RowCount;
                        Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[1].Value = contract_id;
                        //
                        for (int n = 0; n < dataGridView19.RowCount; n++)
                        {
                            if (dataGridView19.Rows[n].Cells[1].Value.ToString() == contract_id) {
                                Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[2].Value = dataGridView19.Rows[n].Cells[3].Value;
                            }
                        }
                        //
                        if ((Frm5.dataGridView1.RowCount - 1) % 2 == 0)
                        {
                            Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[0].Style.BackColor = cl_even;
                            Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[1].Style.BackColor = cl_even;
                            Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[2].Style.BackColor = cl_even;
                        }
                        else
                        {
                            Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[0].Style.BackColor = cl_odd;
                            Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[1].Style.BackColor = cl_odd;
                            Frm5.dataGridView1.Rows[Frm5.dataGridView1.RowCount - 1].Cells[2].Style.BackColor = cl_odd;
                        }
                        //Frm5.comboBox1.Items.Add(contract_id);
                    }
                }
            }
            catch { }
        }

        private void label105_Click(object sender, EventArgs e)
        {
            textBox35.Clear();
        }

        private void label78_Click(object sender, EventArgs e)
        {
            textBox31.Clear();
        }

        private void label75_Click(object sender, EventArgs e)
        {
            textBox30.Clear();
        }

        private void button41_Click(object sender, EventArgs e)
        {
            if (dataGridView12.RowCount != 0)
            {
                int selected = dataGridView10.CurrentCell.RowIndex;
                string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
                string customer = dataGridView10.Rows[selected].Cells[2].Value.ToString();
                //
                selected = dataGridView11.CurrentCell.RowIndex;
                string contract_id = dataGridView11.Rows[selected].Cells[5].Value.ToString();
                string v = dataGridView11.Rows[selected].Cells[0].Value.ToString();
                string cp_date = dataGridView11.Rows[selected].Cells[1].Value.ToString();
                string d_t = dataGridView11.Rows[selected].Cells[2].Value.ToString();
                string c_order_id = dataGridView11.Rows[selected].Cells[3].Value.ToString();
                string c_order_dt = dataGridView11.Rows[selected].Cells[4].Value.ToString();
                string po_order_id = dataGridView11.Rows[selected].Cells[7].Value.ToString();
                string po_order_dt = dataGridView11.Rows[selected].Cells[8].Value.ToString();
                //
                if(contract_id.Length>0)
                {
                    MessageBox.Show("Контракт уже создан", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if(checkBox18.Checked==false && checkBox19.Checked==false && checkBox20.Checked==false)
                    {
                        MessageBox.Show("Please check the terms of payment", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        if (checkBox16.Checked == false && checkBox17.Checked == false)
                        {
                            MessageBox.Show("Please check the terms of delivery", "Сообщение",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            if (textBox40.Text.ToString().Length <= 0 || textBox43.Text.ToString().Length <= 0 || textBox44.Text.ToString().Length <= 0)
                            {
                                MessageBox.Show("Please fill the payment details", "Сообщение",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                if (d_t != "AIR & RW")
                                {
                                    double amount = 0;
                                    for (int i = 0; i < dataGridView12.RowCount - 1; i++)
                                    {
                                        if (d_t == "RW")
                                        {
                                            amount = amount + Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                                        }
                                        if (d_t == "AIR")
                                        {
                                            amount = amount + Math.Round(Convert.ToDouble(dataGridView12.Rows[i].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                                        }
                                    }
                                    createContract(cp_id, cp_date, d_t, v, customer, amount, c_order_id, c_order_dt, po_order_id, po_order_dt);
                                }
                                else
                                {
                                    MessageBox.Show("Контракт может быть составлен либо с ценами RW либо AIR", "Сообщение",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                    }
                }
            }
        }

        public void GotoPO(string cp_id, string customer, string v, string cp_date, string c_order_id, string c_order_dt, string ct_id, string ct_dt, string project)
        {
            // Appearance
            textBox1.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            dateTimePicker1.Enabled = true;
            textBox1.Clear();
            textBox2.Clear();
            textBox5.Clear();
            textBox6.Clear();
            richTextBox1.Clear();
            dataGridView1.RowCount = 0;
            //dataGridView1.AllowUserToAddRows = true;
            //dataGridView1.RowCount = 1;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = false;
            button14.Enabled = true;
            button13.Enabled = false;
            button27.Enabled = true;
            button28.Enabled = true;
            comboBox5.Enabled = true;
            menuStrip1.Items[0].Enabled = false;
            // Funcs
            textBox1.Text = c_order_id;
            string dt_d = c_order_dt.Substring(0, 2);
            string dt_m = c_order_dt.Substring(3, 2);
            string dt_y = c_order_dt.Substring(6, 4);
            dateTimePicker1.Value = new DateTime(Convert.ToInt32(dt_y), Convert.ToInt32(dt_m), Convert.ToInt32(dt_d));
            //
            textBox6.Text = ct_id;
            dt_d = ct_dt.Substring(0, 2);
            dt_m = ct_dt.Substring(3, 2);
            dt_y = ct_dt.Substring(6, 4);
            dateTimePicker8.Value = new DateTime(Convert.ToInt32(dt_y), Convert.ToInt32(dt_m), Convert.ToInt32(dt_d));
            //
            textBox5.Text = cp_id;
            dt_d = cp_date.Substring(0, 2);
            dt_m = cp_date.Substring(3, 2);
            dt_y = cp_date.Substring(6, 4);
            dateTimePicker7.Value = new DateTime(Convert.ToInt32(dt_y), Convert.ToInt32(dt_m), Convert.ToInt32(dt_d));
            comboBox5.Text = v;
            textBox13.Text = customer;
            textBox36.Text = project;
            //
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                int k = dataGridView1.Rows.IndexOf(row);
                if (k % 2 == 0)
                {
                    row.DefaultCellStyle.BackColor = cl_even;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = cl_odd;
                }
                row.HeaderCell.Value = "";
            }

            int indexOfTotalRow = dataGridView1.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[k].Cells[3].Value);
            }
            dataGridView1.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
            dataGridView1.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
            dataGridView1.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.FirstDisplayedScrollingRowIndex + 1;
            // 
            panel_menu_po.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_po.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
        }
        // axad 22.07 start
        public void createContract(string cp_id, string cp_date, string delivery_type, string v, string customer, double amount, string c_order_id, string c_order_dt, string po_order_id, string po_order_dt)
        {
            string contract_date = "";
            string id_contract = "";
            string item_code = "";
            string name = "";
            string quantity = "";
            string unit_price = "";
            string amount_price = "";
            string delivery_point_air = "";
            string delivery_point_rw = "";
            string pay_before_percent = "";
            string pay_before_period = "";
            string delivery_period = "";
            string terms_of_payment = "";
            string terms_of_delivery = "";
            //
            pay_before_percent = textBox44.Text.ToString();
            pay_before_period = textBox43.Text.ToString();
            delivery_period = textBox40.Text.ToString();
            if(checkBox20.Checked==true)
            {
                terms_of_payment = "1";
            }
            if (checkBox19.Checked == true)
            {
                terms_of_payment = "2";
            }
            if (checkBox18.Checked == true)
            {
                terms_of_payment = "3";
            }
            if (checkBox17.Checked == true)
            {
                terms_of_delivery = "1";
            }
            if (checkBox16.Checked == true)
            {
                terms_of_delivery = "2";
            }

            string year = Convert.ToString(dateTimePicker12.Value.Year);
            string month = Convert.ToString(dateTimePicker12.Value.Month);
            string day = Convert.ToString(dateTimePicker12.Value.Day);
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            contract_date = day + "." + month + "." + year;

            delivery_point_air = textBox21.Text;
            delivery_point_rw = textBox22.Text;

            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            SqlCommand command1 = new SqlCommand();
            command1.Connection = connection;
            command1.CommandType = CommandType.Text;
            command1.CommandText = "select count(*) from contracts where date=@date";
            command1.Parameters.AddWithValue("@date", contract_date);
            int count = Convert.ToInt32(command1.ExecuteScalar());
            //
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "select count(*) from contracts";
            int countAll = Convert.ToInt32(command.ExecuteScalar());

            //id_contract = "LSDZ" + date.ToString("yyyyMMdd") + "-" + (countAll + 1).ToString() + "-" + (count+1).ToString();
            id_contract = "LSDZ" + (year + month + day) + "-" + (countAll + 1).ToString() + "-" + (count + 1).ToString();
            //
            command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            command.CommandText = "UPDATE com_proposal SET ct_id='" + id_contract + "', ct_date='" + contract_date + "' WHERE cp_id='" + cp_id + "' and version=" + v;
            command.ExecuteNonQuery();

            SqlCommand command2 = new SqlCommand();
            command2.Connection = connection;
            command2.CommandType = CommandType.Text;
            command2.CommandText = "INSERT INTO contracts(id_contract, date, delivery_type, id_cp, cp_date, client_order_id, client_order_id_date, id_order, id_order_date, delivery_point_air, cp_version,  customer, amount, version, delivery_point_rw, pay_before_percent, pay_before_period, delivery_period, terms_of_payment, terms_of_delivery) VALUES(@idcont, @date, @delivery, @idcp, @cpdate, @clidorder, @clientorderdate, @idorder, @idorderdate, @dpoint_air, @v, @c, @a, @version, @dpoint_rw, @paybefore, @paybefore_period, @delivery_period, @termspay, @termsdel)";
            command2.Parameters.AddWithValue("@idcont", id_contract);
            command2.Parameters.AddWithValue("@date", contract_date);
            command2.Parameters.AddWithValue("@delivery", delivery_type);
            command2.Parameters.AddWithValue("@idcp", cp_id);
            command2.Parameters.AddWithValue("@cpdate", cp_date);
            command2.Parameters.AddWithValue("@clidorder", c_order_id);
            command2.Parameters.AddWithValue("@clientorderdate", c_order_dt);
            command2.Parameters.AddWithValue("@idorder", po_order_id);
            command2.Parameters.AddWithValue("@idorderdate", po_order_dt);
            command2.Parameters.AddWithValue("@dpoint_air", delivery_point_air);
            command2.Parameters.AddWithValue("@v", v);
            command2.Parameters.AddWithValue("@c", customer);
            command2.Parameters.AddWithValue("@a", amount);
            command2.Parameters.AddWithValue("@version", "1");
            command2.Parameters.AddWithValue("@dpoint_rw", delivery_point_rw);
            command2.Parameters.AddWithValue("@paybefore", pay_before_percent);
            command2.Parameters.AddWithValue("@paybefore_period", pay_before_period);
            command2.Parameters.AddWithValue("@delivery_period", delivery_period);
            command2.Parameters.AddWithValue("@termspay", terms_of_payment);
            command2.Parameters.AddWithValue("@termsdel", terms_of_delivery);

            command2.ExecuteNonQuery();

            int i = 0;
            for (i = 0; i < dataGridView12.Rows.Count-1; i++)
            {
                try { item_code = dataGridView12.Rows[i].Cells[2].Value.ToString(); } catch { item_code = ""; }
                try { name = dataGridView12.Rows[i].Cells[3].Value.ToString(); } catch { name = ""; }
                try { quantity = dataGridView12.Rows[i].Cells[5].Value.ToString(); } catch { quantity = ""; }
                if (delivery_type == "AIR")
                {
                    try { unit_price = dataGridView12.Rows[i].Cells[20].Value.ToString(); } catch { unit_price = ""; }
                    try { amount_price = dataGridView12.Rows[i].Cells[22].Value.ToString(); } catch { amount_price = ""; }
                }
                if (delivery_type == "RW")
                {
                    try { unit_price = dataGridView12.Rows[i].Cells[19].Value.ToString(); } catch { unit_price = ""; }
                    try { amount_price = dataGridView12.Rows[i].Cells[21].Value.ToString(); } catch { amount_price = ""; }
                }

                SqlCommand command3 = new SqlCommand();
                command3.Connection = connection;
                command3.CommandType = CommandType.Text;
                command3.CommandText = "INSERT INTO items_in_contract(id_contract, item_code, name, quantity, unit_price, amount_price, version) VALUES(@idcont, @itemcode, @name, @quantity, @unitprice, @amountprice, @v)";
                command3.Parameters.AddWithValue("@idcont", id_contract);
                command3.Parameters.AddWithValue("@itemcode", item_code);
                command3.Parameters.AddWithValue("@name", name);
                command3.Parameters.AddWithValue("@quantity", quantity);
                command3.Parameters.AddWithValue("@unitprice", unit_price);
                command3.Parameters.AddWithValue("@amountprice", amount_price);
                command3.Parameters.AddWithValue("@v", '1');
                command3.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("Контракт с ID  " + id_contract + "  успешно создан на основе комерческого предложения  " + cp_id + "  версии  " + v, "Done",
            MessageBoxButtons.OK, MessageBoxIcon.Information);

            string delivery_point = "";
            if(delivery_type=="AIR")
            {
                delivery_point = delivery_point_air;
            }
            if(delivery_type=="RW")
            {
                delivery_point = delivery_point_rw;
            }
            generateContractFromCP(id_contract, contract_date, delivery_type, delivery_point, customer);
        }
        // axad 22.07 end

        private void label66_Click_1(object sender, EventArgs e)
        {
            textBox26.Clear();
        }

        private void label70_Click(object sender, EventArgs e)
        {
            textBox28.Clear();
        }

        private void label73_Click(object sender, EventArgs e)
        {
            textBox29.Clear();
        }

        private void button39_Click_1(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                GenerateWhiteID(Convert.ToInt32(numericUpDown5.Value.ToString()), "gen", textBox27.Text, textBox33.Text);
            }
            else
            {
                GenerateWhiteID(Convert.ToInt32(numericUpDown5.Value.ToString()), "print", textBox27.Text, textBox33.Text);
            }
        }

        private void dataGridView19_SelectionChanged(object sender, EventArgs e)
        {
            int selected = -1;
            try {
                selected = dataGridView19.CurrentCell.RowIndex;
                textBox27.Text = dataGridView19.Rows[selected].Cells[1].Value.ToString();
                textBox33.Text = dataGridView19.Rows[selected].Cells[3].Value.ToString();
            }
            catch { }
            
        }

        private void checkBox7_Click(object sender, EventArgs e)
        {
            checkBox8.Checked = false;
        }

        private void checkBox8_Click(object sender, EventArgs e)
        {
            checkBox7.Checked = false;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            if (dataGridView11.RowCount != 0) {
                int selected = dataGridView10.CurrentCell.RowIndex;
                string cp_id = dataGridView10.Rows[selected].Cells[1].Value.ToString();
                string customer = dataGridView10.Rows[selected].Cells[2].Value.ToString();
                string project = dataGridView10.Rows[selected].Cells[3].Value.ToString();
                //
                selected = dataGridView11.CurrentCell.RowIndex;
                string v = dataGridView11.Rows[selected].Cells[0].Value.ToString();
                string cp_date = dataGridView11.Rows[selected].Cells[1].Value.ToString();
                string d_t = dataGridView11.Rows[selected].Cells[2].Value.ToString();
                string c_order_id = dataGridView11.Rows[selected].Cells[3].Value.ToString();
                string c_order_dt = dataGridView11.Rows[selected].Cells[4].Value.ToString();
                string ct_id = dataGridView11.Rows[selected].Cells[5].Value.ToString();
                string ct_dt = dataGridView11.Rows[selected].Cells[6].Value.ToString();
                if (ct_id != "")
                {
                    ShowPanels("Заказ");
                    textBox1.Clear();
                    textBox2.Clear();
                    Size ns = new Size();
                    ns.Width = panel7.Width - 30;
                    ns.Height = panel7.Height - ((panel7.Height / 2) - (panel7.Height / 11));
                    //richTextBox1.Size = ns;
                    Point np = new Point();
                    np.Y = richTextBox1.Location.Y + richTextBox1.Size.Height + 15;
                    np.X = 15;
                    //button3.Location = np;
                    checkBox1.Checked = false;
                    checkBox2.Checked = true;
                    // Init gridview
                    dataGridView1.RowCount = 0;
                    dataGridView1.ColumnCount = 8;
                    //
                    dataGridView1.Columns[0].HeaderText = "#";
                    dataGridView1.Columns[1].HeaderText = "Код продукта";
                    dataGridView1.Columns[2].HeaderText = "Название продукта";
                    dataGridView1.Columns[3].HeaderText = "Количество";
                    dataGridView1.Columns[4].HeaderText = "";
                    dataGridView1.Columns[5].HeaderText = "Не хватает";
                    dataGridView1.Columns[6].HeaderText = "На складе";
                    dataGridView1.Columns[7].HeaderText = "Заметки";
                    //
                    dataGridView1.Columns[0].Width = 50;
                    dataGridView1.Columns[1].Width = 120;
                    dataGridView1.Columns[2].Width = 200;
                    dataGridView1.Columns[3].Width = 80;
                    dataGridView1.Columns[4].Width = 40;
                    dataGridView1.Columns[5].Width = 90;
                    dataGridView1.Columns[6].Width = 90;
                    dataGridView1.Columns[7].Width = (dataGridView1.Size.Width - (dataGridView1.Columns[0].Width + dataGridView1.Columns[1].Width + dataGridView1.Columns[2].Width + dataGridView1.Columns[3].Width + dataGridView1.Columns[4].Width + dataGridView1.Columns[5].Width + dataGridView1.Columns[6].Width)) - 100;
                    // // //
                    dataGridView4.RowCount = 0;
                    dataGridView4.ColumnCount = 10;
                    //
                    dataGridView4.Columns[0].HeaderText = "#";
                    dataGridView4.Columns[1].HeaderText = "PO от";
                    dataGridView4.Columns[2].HeaderText = "PO ID";
                    dataGridView4.Columns[3].HeaderText = "PO дата";
                    dataGridView4.Columns[4].HeaderText = "Order ID";
                    dataGridView4.Columns[5].HeaderText = "Order дата";
                    dataGridView4.Columns[6].HeaderText = "CP ID";
                    dataGridView4.Columns[7].HeaderText = "CP дата";
                    dataGridView4.Columns[8].HeaderText = "CT ID";
                    dataGridView4.Columns[9].HeaderText = "CT дата";
                    //
                    //dataGridView4.Columns[0].Width = 90;
                    //dataGridView4.Columns[1].Width = 80;
                    //dataGridView4.Columns[2].Width = 80;
                    //dataGridView4.Columns[3].Width = 90;
                    //dataGridView4.Columns[4].Width = 80;
                    //dataGridView4.Columns[5].Width = 80;
                    //dataGridView4.Columns[6].Width = 130;
                    // // //
                    dataGridView5.RowCount = 0;
                    dataGridView5.ColumnCount = 2;
                    //
                    dataGridView5.Columns[0].HeaderText = "Версия";
                    dataGridView5.Columns[1].HeaderText = "Дата";
                    //
                    dataGridView5.Columns[0].Width = 70;
                    dataGridView5.Columns[1].Width = 70;
                    // // //
                    dataGridView6.RowCount = 0;
                    dataGridView6.ColumnCount = 8;
                    //
                    dataGridView6.Columns[0].HeaderText = "#";
                    dataGridView6.Columns[1].HeaderText = "Код продукта";
                    dataGridView6.Columns[2].HeaderText = "Название продукта";
                    dataGridView6.Columns[3].HeaderText = "Кол-во";
                    dataGridView6.Columns[4].HeaderText = "";
                    dataGridView6.Columns[5].HeaderText = "Не хватает";
                    dataGridView6.Columns[6].HeaderText = "На складе";
                    dataGridView6.Columns[7].HeaderText = "Заметки";
                    //
                    dataGridView6.Columns[0].Width = 50;
                    dataGridView6.Columns[1].Width = 130;
                    dataGridView6.Columns[2].Width = 200;
                    dataGridView6.Columns[3].Width = 60;
                    dataGridView6.Columns[4].Width = 30;
                    dataGridView6.Columns[5].Width = 60;
                    dataGridView6.Columns[6].Width = 60;
                    int sz = dataGridView6.Columns[0].Width + dataGridView6.Columns[1].Width + dataGridView6.Columns[2].Width + dataGridView6.Columns[3].Width + dataGridView6.Columns[4].Width + dataGridView6.Columns[5].Width + dataGridView6.Columns[6].Width;
                    dataGridView6.Columns[7].Width = dataGridView6.Size.Width - sz - 60;
                    //
                    GotoPO(cp_id, customer, v, cp_date, c_order_id, c_order_dt, ct_id, ct_dt, project);
                }
                else
                {
                    MessageBox.Show("Невозможно создать заказа в LSIS (Production Order) так как на данное комерческое предложение не создан контракт.", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            
        }

        private void dataGridView15_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            LoadCTListToCollect();
        }

        private void button40_Click_1(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                GeneratePaletID(Convert.ToInt32(numericUpDown6.Value.ToString()), "gen", textBox27.Text, textBox33.Text);
            }
            else
            {
                GeneratePaletID(Convert.ToInt32(numericUpDown6.Value.ToString()), "print", textBox27.Text, textBox33.Text);
            }
        }

        private void checkBox9_Click(object sender, EventArgs e)
        {
            checkBox9.Checked = true;
            checkBox10.Checked = false;
        }

        private void checkBox10_Click(object sender, EventArgs e)
        {
            checkBox10.Checked = true;
            checkBox9.Checked = false;
        }

        private void button42_Click(object sender, EventArgs e)
        {
            saveCollectedContracts();
        }

        public void saveCollectedContracts()
        {
            bool check = true;
            for (int k = 0; k < dataGridView16.Rows.Count; k++)
            {
                int needed = 0;
                try { needed = Convert.ToInt32(dataGridView16.Rows[k].Cells[5].Value.ToString()); } catch { needed = 0; }
                if (needed > 0)
                {
                    check = false;
                }
            }
            if (check == true)
            {
                for (int i = 0; i < dataGridView19.RowCount; i++)
                {
                    string id_contract = dataGridView19.Rows[i].Cells[1].Value.ToString();
                    string date = "";
                    string delivery_type = "";
                    string id_cp = "";
                    string cp_version = "";
                    string client_order_id = "";
                    string id_order = "";
                    string delivery_point = "";
                    string customer = "";
                    double amount = 0;
                    string cp_date = "";
                    string client_order_id_date = "";
                    string id_order_date = "";
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        // axad 22.07 start
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, delivery_type, id_cp, client_order_id, id_order, delivery_point_air, delivery_point_rw, cp_version, customer, amount, cp_date, client_order_id_date, id_order_date FROM contracts WHERE id_contract='" + id_contract + "'", connection))
                        {
                            // axad 22.07 end
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    date = reader.GetValue(1).ToString();
                                    delivery_type = reader.GetValue(2).ToString();
                                    id_cp = reader.GetValue(3).ToString();
                                    client_order_id = reader.GetValue(4).ToString();
                                    id_order = reader.GetValue(5).ToString();
                                    if(reader.GetValue(2).ToString()=="AIR")
                                    {
                                        delivery_point = reader.GetValue(6).ToString();
                                    }
                                    if (reader.GetValue(2).ToString() == "rw")
                                    {
                                        delivery_point = reader.GetValue(7).ToString();
                                    }
                                    cp_version = reader.GetValue(8).ToString();
                                    customer = reader.GetValue(9).ToString();
                                    amount = Math.Round(Convert.ToDouble(reader.GetValue(10).ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                                    cp_date = reader.GetValue(11).ToString();
                                    client_order_id_date = reader.GetValue(12).ToString();
                                    id_order_date = reader.GetValue(13).ToString();
                                }
                            }
                        }
                        connection.Close();
                    }
                    SqlConnection connection1 = new SqlConnection(conString);
                    connection1.Open();

                    SqlCommand command1 = new SqlCommand();
                    command1.Connection = connection1;
                    command1.CommandType = CommandType.Text;
                    command1.CommandText = "INSERT INTO collected_contracts(id_contract, date, delivery_type, id_cp, client_order_id, id_order, delivery_point, cp_version, cp_date, client_order_id_date, id_order_date, amount, customer) VALUES(@idcontract, @date, @deliverytype, @idcp, @clientorderid, @idorder, @deliverypoint, @cpversion, @cpdate, @cloriddate, @idorderdate, @amount, @customer)";
                    command1.Parameters.AddWithValue("@idcontract", id_contract);
                    command1.Parameters.AddWithValue("@date", date);
                    command1.Parameters.AddWithValue("@deliverytype", delivery_type);
                    command1.Parameters.AddWithValue("@idcp", id_cp);
                    command1.Parameters.AddWithValue("@clientorderid", client_order_id);
                    command1.Parameters.AddWithValue("@idorder", id_order);
                    command1.Parameters.AddWithValue("@deliverypoint", delivery_point);
                    command1.Parameters.AddWithValue("@cpversion", cp_version);
                    command1.Parameters.AddWithValue("@cpdate", cp_date);
                    command1.Parameters.AddWithValue("@cloriddate", client_order_id_date);
                    command1.Parameters.AddWithValue("@idorderdate", id_order_date);
                    command1.Parameters.AddWithValue("@amount", amount);
                    command1.Parameters.AddWithValue("@customer", customer);
                    command1.ExecuteNonQuery();

                }
            }
            else
            {
                MessageBox.Show("Заказ собран не полностью.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void GetContractListToGenerate(string argument, int type)
        {
            // search argument cp_id
            if (type == 0)
            {
                if (argument == "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, delivery_type, delivery_point FROM collected_contracts", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.date = reader.GetValue(1).ToString();
                                    item.delivery_type = reader.GetValue(2).ToString();
                                    item.delivery_point = reader.GetValue(3).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView20.RowCount = 0;
                    dataGridView20.RowCount = contracts.Count;
                    dataGridView20.RowHeadersWidth = 35;
                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView20.Rows[i].Cells[0].Value = i + 1;
                        dataGridView20.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView20.Rows[i].Cells[2].Value = contracts[i].date;
                        if (i % 2 == 0)
                        {
                            dataGridView20.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView20.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView20.Rows[i].Cells[2].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView20.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView20.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView20.Rows[i].Cells[2].Style.BackColor = cl_odd;
                        }
                    }
                }
                //
                if (argument != "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, delivery_type, delivery_point FROM collected_contracts WHERE id_contract ='" + argument + "'", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.date = reader.GetValue(1).ToString();
                                    item.delivery_type = reader.GetValue(2).ToString();
                                    item.delivery_point = reader.GetValue(3).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView20.RowCount = 0;
                    dataGridView20.RowCount = contracts.Count;
                    dataGridView20.RowHeadersWidth = 35;
                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView20.Rows[i].Cells[0].Value = i + 1;
                        dataGridView20.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView20.Rows[i].Cells[2].Value = contracts[i].date;
                        if (i % 2 == 0)
                        {
                            dataGridView20.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView20.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView20.Rows[i].Cells[2].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView20.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView20.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView20.Rows[i].Cells[2].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
            // search argument date
            if (type == 1)
            {
                //
                if (argument != "")
                {
                    List<Contracts> contracts = new List<Contracts>();
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand("SELECT id_contract, date, delivery_type, delivery_point FROM collected_contracts WHERE date='" + argument + "'", connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Contracts item = new Contracts();
                                    item.id_contract = reader.GetValue(0).ToString();
                                    item.date = reader.GetValue(1).ToString();
                                    item.delivery_type = reader.GetValue(2).ToString();
                                    item.delivery_point = reader.GetValue(3).ToString();
                                    contracts.Add(item);
                                }
                            }
                        }
                        connection.Close();
                    }
                    dataGridView20.RowCount = 0;
                    dataGridView20.RowCount = contracts.Count;
                    dataGridView20.RowHeadersWidth = 35;
                    for (int i = 0; i < contracts.Count; i++)
                    {
                        dataGridView20.Rows[i].Cells[0].Value = i + 1;
                        dataGridView20.Rows[i].Cells[1].Value = contracts[i].id_contract;
                        dataGridView20.Rows[i].Cells[2].Value = contracts[i].date;

                        if (i % 2 == 0)
                        {
                            dataGridView20.Rows[i].Cells[0].Style.BackColor = cl_even;
                            dataGridView20.Rows[i].Cells[1].Style.BackColor = cl_even;
                            dataGridView20.Rows[i].Cells[2].Style.BackColor = cl_even;
                        }
                        else
                        {
                            dataGridView20.Rows[i].Cells[0].Style.BackColor = cl_odd;
                            dataGridView20.Rows[i].Cells[1].Style.BackColor = cl_odd;
                            dataGridView20.Rows[i].Cells[2].Style.BackColor = cl_odd;
                        }
                    }
                }
            }
        }

        public class CollectedContractItems
        {
            public string idContract { get; set; }
            public string idOrderClient { get; set; }
            public string idOrderClientDate { get; set; }
            public string itemCode { get; set; }
            public string itemName { get; set; }
            public string hsCode { get; set; }
            public string quantity { get; set; }
            public string unitPrice { get; set; }
            public string amountPrice { get; set; }
            public string netWeight { get; set; }
            public string grossWeight { get; set; }
            public string boxNo { get; set; }
            public string poleta { get; set; }
        }

        public void getCollectedContractItems(string idContract)
        {
            List<CollectedContractItems> contract_items = new List<CollectedContractItems>();
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("select collected_contracts.id_contract, collected_contracts.client_order_id, collected_contracts.client_order_id_date, collected_contract_items.item_code, collected_contract_items.name, items_in_cp.hs_code, collected_contract_items.quantity, items_in_cp.unit_price, items_in_cp.amount_price, items_in_cp.weight_of_item, collected_contract_items.white_id, collected_contract_items.no_poleta  from collected_contracts join collected_contract_items on collected_contract_items.id_contract=collected_contracts.id_contract join items_in_cp on items_in_cp.cp_id=collected_contracts.id_cp and items_in_cp.version=collected_contracts.cp_version and items_in_cp.part_code=collected_contract_items.item_code where collected_contracts.id_contract=@idcontract order by collected_contract_items.white_id", connection))
                {
                    command.Parameters.AddWithValue("@idcontract", idContract);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CollectedContractItems item = new CollectedContractItems();
                            item.idContract = reader.GetValue(0).ToString();
                            item.idOrderClient = reader.GetValue(1).ToString();
                            item.idOrderClientDate = reader.GetValue(2).ToString();
                            item.itemCode = reader.GetValue(3).ToString();
                            item.itemName = reader.GetValue(4).ToString();
                            item.hsCode = reader.GetValue(5).ToString();
                            item.quantity = reader.GetValue(6).ToString();
                            item.unitPrice = reader.GetValue(7).ToString();
                            item.amountPrice = reader.GetValue(8).ToString();
                            item.netWeight = Math.Round(Math.Round(Convert.ToDouble(reader.GetValue(9).ToString(), new System.Globalization.CultureInfo("en-US")) , 2, MidpointRounding.ToEven)* Math.Round(Convert.ToDouble(reader.GetValue(6).ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven), 2, MidpointRounding.ToEven).ToString();
                            item.grossWeight = "";
                            item.boxNo = reader.GetValue(10).ToString();
                            item.poleta = reader.GetValue(11).ToString();
                            contract_items.Add(item);
                        }
                    }
                }
                connection.Close();
            }
            int incr = dataGridView21.RowCount;
            int k = 0;
            dataGridView21.RowCount = incr + contract_items.Count;
            for (int i = incr; i < dataGridView21.RowCount; i++)
            {
                dataGridView21.Rows[i].Cells[0].Value = i + 1;
                dataGridView21.Rows[i].Cells[1].Value = contract_items[k].idContract;
                dataGridView21.Rows[i].Cells[2].Value = contract_items[k].idOrderClient;
                dataGridView21.Rows[i].Cells[3].Value = contract_items[k].idOrderClientDate;
                dataGridView21.Rows[i].Cells[4].Value = contract_items[k].itemCode;
                dataGridView21.Rows[i].Cells[5].Value = contract_items[k].itemName;
                dataGridView21.Rows[i].Cells[6].Value = contract_items[k].hsCode;
                dataGridView21.Rows[i].Cells[7].Value = contract_items[k].quantity;
                dataGridView21.Rows[i].Cells[8].Value = contract_items[k].unitPrice;
                dataGridView21.Rows[i].Cells[9].Value = contract_items[k].amountPrice;
                dataGridView21.Rows[i].Cells[10].Value = contract_items[k].netWeight;
                dataGridView21.Rows[i].Cells[11].Value = contract_items[k].grossWeight;
                dataGridView21.Rows[i].Cells[12].Value = contract_items[k].boxNo;
                dataGridView21.Rows[i].Cells[13].Value = contract_items[k].poleta;
                k++;
                if (i % 2 == 0)
                {
                    dataGridView21.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[5].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[6].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[7].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[8].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[9].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[10].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[11].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[12].Style.BackColor = cl_even;
                    dataGridView21.Rows[i].Cells[13].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView21.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[5].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[6].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[7].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[8].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[9].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[10].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[11].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[12].Style.BackColor = cl_odd;
                    dataGridView21.Rows[i].Cells[13].Style.BackColor = cl_odd;
                }
            }
        }

        public class Polets
        {
            public string poletNo { get; set; }
        }

        public void MakePLDoc()
        {
            DateTime dateTime = DateTime.Now;
            string year = dateTime.Year.ToString();
            string month = dateTime.Month.ToString();
            string day = dateTime.Day.ToString();
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string today = day + "." + month + "." + year;

            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");

            worksheet.Row(1).Height = 60;
            worksheet.Row(2).Height = 40;
            worksheet.Row(18).Height = 35;
            worksheet.Column(1).Width = 6;
            worksheet.Column(2).Width = 6;
            worksheet.Column(3).Width = 5;
            worksheet.Column(4).Width = 3;
            worksheet.Column(5).Width = 8;
            worksheet.Column(6).Width = 23;
            worksheet.Column(7).Width = 6;
            worksheet.Column(8).Width = 6;
            worksheet.Column(9).Width = 5;
            worksheet.Column(10).Width = 5;
            worksheet.Column(11).Width = 6;
            worksheet.Column(12).Width = 5;
            worksheet.Column(13).Width = 5;

            string year_arriving = Convert.ToString(dateTimePicker16.Value.Year);
            string month_arriving = Convert.ToString(dateTimePicker16.Value.Month);
            string day_arriving = Convert.ToString(dateTimePicker16.Value.Day);
            if (month_arriving.Length == 1) { month_arriving = "0" + month_arriving; }
            if (day_arriving.Length == 1) { day_arriving = "0" + day_arriving; }
            string arriving_date = day_arriving + "." + month_arriving + "." + year_arriving;
            worksheet.Cells[2, 1].Value = "Packing List/ Упаковочный лист";
            worksheet.Cells[3, 1].Value = "1. Отправитель/Экспортер (Shipper/Exporter)";
            worksheet.Cells[4, 1].Value = textBoxExporterCompany.Text;
            worksheet.Cells[5, 1].Value = textBoxExporterAddress.Text;
            worksheet.Cells[6, 1].Value = textBoxExporterContact.Text;
            worksheet.Cells[7, 1].Value = "2. Получатель (Consignee)";
            worksheet.Cells[8, 1].Value = comboBox10.GetItemText(comboBox10.SelectedItem);
            worksheet.Cells[9, 1].Value = comboBox11.GetItemText(comboBox11.SelectedItem);
            worksheet.Cells[10, 1].Value = comboBox12.GetItemText(comboBox12.SelectedItem);
            worksheet.Cells[12, 1].Value = "3. Извещаемая сторона (Notify party)";
            worksheet.Cells[13, 1].Value = textBoxNotifyCompany.Text;
            worksheet.Cells[14, 1].Value = textBoxNotifyAdress.Text;
            worksheet.Cells[15, 1].Value = "4. Порт загрузки (Loading port)";
            worksheet.Cells[16, 1].Value = textBoxLoadingPort.Text;
            worksheet.Cells[15, 6].Value = "5. Порт разгрузки (Port of Discharging)";
            worksheet.Cells[16, 6].Value = comboBox13.GetItemText(comboBox13.SelectedItem);
            worksheet.Cells[17, 1].Value = "6. Пункт назначения(Final Destination)";
            worksheet.Cells[18, 1].Value = textBoxFinalDestination.Text;
            worksheet.Cells[17, 6].Value = "7. Перевозчик (Carrier)";
            worksheet.Cells[18, 6].Value = textBoxCarrier.Text;
            worksheet.Cells[19, 1].Value = "8. Номер рейса(Voyage No.)";
            worksheet.Cells[20, 1].Value = textBoxVoyageNo.Text;
            worksheet.Cells[19, 6].Value = "9. Примерная дата отгрузки (Sailing on or about)";
            worksheet.Cells[20, 6].Value = arriving_date;

            worksheet.Cells[3, 8].Value = "10. Номер и дата инвойса (No. & Date of invoice)";
            worksheet.Cells[4, 8].Value = "";
            worksheet.Cells[4, 11].Value = today;
            worksheet.Cells[7, 8].Value = "11. Примечание (Remarks)";
            worksheet.Cells[7, 11].Value = "Order ";
            worksheet.Cells[17, 8].Value = "Country of Origin: " + textBoxOrigin.Text;
            worksheet.Cells[18, 8].Value = "Terms of delivery : " + textBoxConditions.Text;
            using (ExcelRange rng = worksheet.Cells[1, 1, 1, 13])
            {
                rng.Merge = true;
            }
            using (System.Drawing.Image logo = System.Drawing.Image.FromFile(app_dir_temp + "Logo.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("My Logo", logo);
                excelImage.SetPosition(0, 0, 5, 40);
                excelImage.SetSize(20);
            }
            using (ExcelRange rng = worksheet.Cells[2, 1, 2, 13])
            {
                rng.Merge = true;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Font.Size = 15;
                rng.Style.Font.Bold = true;
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[3, 1, 3, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[4, 1, 4, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[5, 1, 5, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[6, 1, 6, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[7, 1, 7, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[8, 1, 8, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[9, 1, 9, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[10, 1, 10, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[11, 1, 11, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[12, 1, 12, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[13, 1, 13, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[14, 1, 14, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[15, 1, 15, 5])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[15, 6, 15, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[16, 1, 16, 5])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[16, 6, 16, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[17, 1, 17, 5])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[17, 6, 17, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[18, 1, 18, 5])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[18, 6, 18, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[19, 1, 19, 5])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[19, 6, 19, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[20, 1, 20, 5])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[20, 6, 20, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[3, 7, 20, 7])
            {
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[3, 8, 3, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[4, 8, 4, 10])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[4, 11, 4, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[5, 8, 5, 13])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[6, 8, 6, 13])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[7, 8, 7, 10])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[7, 11, 7, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            for (int i = 8; i < 17; i++)
            {
                using (ExcelRange rng = worksheet.Cells[i, 8, i, 10])
                {
                    rng.Merge = true;
                }
                using (ExcelRange rng = worksheet.Cells[i, 11, i, 13])
                {
                    rng.Merge = true;
                }
            }
            using (ExcelRange rng = worksheet.Cells[17, 8, 17, 13])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[18, 8, 18, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[19, 8, 19, 13])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[20, 8, 20, 13])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            worksheet.Cells[21, 1].Value = "12. Маркировка & NO of PKGS";
            worksheet.Cells[21, 1, 21, 3].Merge = true;

            worksheet.Cells[21, 4].Value = "Items Code";
            worksheet.Cells[21, 4, 21, 5].Merge = true;

            worksheet.Cells[21, 6].Value = "14. Описание торвара(Description of Goods)";

            worksheet.Cells[21, 7].Value = "15. Quantity (EA)";

            worksheet.Cells[21, 8].Value = "16. Net Weight (Kg)";

            worksheet.Cells[21, 9].Value = "17. Gross Weight(Kg)";

            worksheet.Cells[21, 10].Value = "18. Box Q-ty";

            worksheet.Cells[21, 11].Value = "19 Box No";

            worksheet.Cells[21, 12].Value = "Pallet No";

            worksheet.Cells[21, 13].Value = "20. Measurement(CBM)";

            using (ExcelRange rng = worksheet.Cells[21, 1, 21, 13])
            {
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                rng.Style.Font.Name = "Times New Roman";
            }

            string box_no = "";
            int rowCount = 22;
            int subTotalQty = 0;
            int totalQty = 0;
            int totalBoxes = 0;
            double subtotalNetWeight = 0;
            double totalNetWeight = 0;
            int startIndex = 22;

            for (int i = 0; i < dataGridView21.RowCount; i++)
            {
                string box_number = "";
                try { box_number = dataGridView21.Rows[i].Cells[12].Value.ToString(); } catch { box_number = ""; }

                if (box_number != box_no && subTotalQty == 0)
                {
                    box_no = box_number;

                    startIndex = rowCount;
                    worksheet.Cells[rowCount, 1].Value = "Id Contract: " + dataGridView21.Rows[i].Cells[1].Value.ToString();// + "\nId Order: " + dataGridView21.Rows[i].Cells[2].Value.ToString();

                    worksheet.Cells[rowCount, 4].Value = dataGridView21.Rows[i].Cells[4].Value.ToString();
                    worksheet.Cells[rowCount, 4, rowCount, 5].Merge = true;
                    worksheet.Cells[rowCount, 6].Value = dataGridView21.Rows[i].Cells[5].Value.ToString();
                    worksheet.Cells[rowCount, 7].Value = dataGridView21.Rows[i].Cells[7].Value.ToString();
                    worksheet.Cells[rowCount, 8].Value = dataGridView21.Rows[i].Cells[10].Value.ToString();
                    worksheet.Cells[rowCount, 9].Value = "";
                    worksheet.Cells[rowCount, 10].Value = "";

                    worksheet.Cells[rowCount, 11].Value = dataGridView21.Rows[i].Cells[12].Value.ToString();
                    worksheet.Cells[rowCount, 12].Value = dataGridView21.Rows[i].Cells[13].Value.ToString();

                    subtotalNetWeight = subtotalNetWeight + Math.Round(Convert.ToDouble(dataGridView21.Rows[i].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    subTotalQty = subTotalQty + Convert.ToInt32(dataGridView21.Rows[i].Cells[7].Value.ToString());
                    rowCount++;
                }
                else if (box_number == box_no)
                {
                    worksheet.Cells[rowCount, 4].Value = dataGridView21.Rows[i].Cells[4].Value.ToString();
                    worksheet.Cells[rowCount, 4, rowCount, 5].Merge = true;
                    worksheet.Cells[rowCount, 6].Value = dataGridView21.Rows[i].Cells[5].Value.ToString();
                    worksheet.Cells[rowCount, 7].Value = dataGridView21.Rows[i].Cells[7].Value.ToString();
                    worksheet.Cells[rowCount, 8].Value = dataGridView21.Rows[i].Cells[10].Value.ToString();
                    worksheet.Cells[rowCount, 9].Value = "";
                    worksheet.Cells[rowCount, 10].Value = "";

                    subtotalNetWeight = subtotalNetWeight + Math.Round(Convert.ToDouble(dataGridView21.Rows[i].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
            subTotalQty = subTotalQty + Convert.ToInt32(dataGridView21.Rows[i].Cells[7].Value.ToString());
                    rowCount++;
                }
                else if (box_number != box_no && subTotalQty > 0)
                {
                    worksheet.Cells[startIndex, 1, rowCount, 3].Merge = true;
                    worksheet.Cells[startIndex, 11, rowCount, 11].Merge = true;
                    worksheet.Cells[startIndex, 12, rowCount, 12].Merge = true;
                    worksheet.Cells[rowCount, 4, rowCount, 6].Merge = true;
                    worksheet.Cells[rowCount, 4].Value = "SUBTOTAL FOR #" + box_no.ToString();
                    worksheet.Cells[rowCount, 7].Value = subTotalQty + " EA";
                    worksheet.Cells[rowCount, 8].Value = subtotalNetWeight + " KG";
                    using (ExcelRange rng = worksheet.Cells[rowCount, 4, rowCount, 10])
                    {
                        rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                    totalNetWeight = totalNetWeight + subtotalNetWeight;
                    totalQty = totalQty + subTotalQty;
                    totalBoxes++;
                    subTotalQty = 0;
                    subtotalNetWeight = 0;

                    rowCount++;

                    startIndex = rowCount;
                    worksheet.Cells[rowCount, 1].Value = "Id Contract: " + dataGridView21.Rows[i].Cells[1].Value.ToString();// + "\nId Order: " + dataGridView21.Rows[i].Cells[2].Value.ToString();

                    worksheet.Cells[rowCount, 4].Value = dataGridView21.Rows[i].Cells[4].Value.ToString();
                    worksheet.Cells[rowCount, 4, rowCount, 5].Merge = true;
                    worksheet.Cells[rowCount, 6].Value = dataGridView21.Rows[i].Cells[5].Value.ToString();
                    worksheet.Cells[rowCount, 7].Value = dataGridView21.Rows[i].Cells[7].Value.ToString();
                    worksheet.Cells[rowCount, 8].Value = dataGridView21.Rows[i].Cells[10].Value.ToString();
                    worksheet.Cells[rowCount, 9].Value = "";
                    worksheet.Cells[rowCount, 10].Value = "";

                    worksheet.Cells[rowCount, 11].Value = dataGridView21.Rows[i].Cells[12].Value.ToString();
                    worksheet.Cells[rowCount, 12].Value = dataGridView21.Rows[i].Cells[13].Value.ToString();

                    subtotalNetWeight = subtotalNetWeight + Math.Round(Convert.ToDouble(dataGridView21.Rows[i].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                    subTotalQty = subTotalQty + Convert.ToInt32(dataGridView21.Rows[i].Cells[7].Value.ToString());
                    rowCount++;
                    box_no = box_number;
                }
            }

            worksheet.Cells[startIndex, 1, rowCount, 3].Merge = true;
            worksheet.Cells[startIndex, 11, rowCount, 11].Merge = true;
            worksheet.Cells[startIndex, 12, rowCount, 12].Merge = true;
            worksheet.Cells[rowCount, 4, rowCount, 6].Merge = true;
            worksheet.Cells[rowCount, 4].Value = "SUBTOTAL FOR #" + box_no.ToString();
            worksheet.Cells[rowCount, 7].Value = subTotalQty + " EA";
            worksheet.Cells[rowCount, 8].Value = subtotalNetWeight + " KG";
            using (ExcelRange rng = worksheet.Cells[rowCount, 4, rowCount, 10])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }
            totalNetWeight = totalNetWeight + subtotalNetWeight;
            totalQty = totalQty + subTotalQty;
            totalBoxes++;
            rowCount++;
            List<Polets> list_polets = new List<Polets>();
            for (int n = 0; n < dataGridView21.RowCount; n++)
            {
                bool check_poleta = true;
                string poleta = "";
                try { poleta = dataGridView21.Rows[n].Cells[13].Value.ToString(); } catch { poleta = ""; }
                for (int m = 0; m < list_polets.Count; m++)
                {
                    if (list_polets[m].poletNo == poleta)
                    {
                        check_poleta = false;
                    }
                }
                if (check_poleta == true)
                {
                    Polets p = new Polets();
                    p.poletNo = poleta;
                    list_polets.Add(p);
                }
            }


            worksheet.Cells[rowCount, 1, rowCount, 6].Merge = true;
            worksheet.Cells[rowCount, 1].Value = "TOTAL: ";
            worksheet.Cells[rowCount, 7].Value = totalQty + " EA";
            worksheet.Cells[rowCount, 8].Value = totalNetWeight + " KG";
            worksheet.Cells[rowCount, 11].Value = totalBoxes + " Boxes";
            worksheet.Cells[rowCount, 12].Value = list_polets.Count + " Polets";
            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 13])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
            }
            using (System.Drawing.Image sign = System.Drawing.Image.FromFile(app_dir_temp + "Sign.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Sign", sign);
                excelImage.SetPosition(rowCount + 3, 20, 6, 0);
                excelImage.SetSize(18);
            }
            using (System.Drawing.Image stamp = System.Drawing.Image.FromFile(app_dir_temp + "Stamp.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Stamp", stamp);
                excelImage.SetPosition(rowCount + 3, 0, 9, 20);
                excelImage.SetSize(35);
            }
            worksheet.Cells[rowCount + 6, 6, rowCount + 6, 11].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[rowCount + 7, 6].Value = " Authorized and signed by Anna Li";
            worksheet.Cells[rowCount + 7, 6, rowCount + 7, 9].Merge = true;
            worksheet.Cells[rowCount + 8, 6].Value = " by Power of Attorney on behalf of ";
            worksheet.Cells[rowCount + 8, 6, rowCount + 8, 9].Merge = true;
            worksheet.Cells[rowCount + 9, 6].Value = " Drone Zone General Manager  ";
            worksheet.Cells[rowCount + 9, 6, rowCount + 9, 9].Merge = true;
            worksheet.Cells[rowCount + 10, 6].Value = " dronezone.anna@gmail.com  ";
            worksheet.Cells[rowCount + 10, 6].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 10, 6, rowCount + 10, 9].Merge = true;

            worksheet.Cells[rowCount + 9, 1].Value = " Customer care  ";
            worksheet.Cells[rowCount + 9, 1, rowCount + 9, 4].Merge = true;
            worksheet.Cells[rowCount + 10, 1].Value = " dronezone.sk@gmail.com  ";
            worksheet.Cells[rowCount + 10, 1].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 10, 1, rowCount + 10, 5].Merge = true;

            worksheet.Cells[rowCount + 7, 10].Value = today;
            worksheet.Cells[rowCount + 7, 10, rowCount + 7, 11].Merge = true;
            using (ExcelRange rng = worksheet.Cells[rowCount + 7, 1, rowCount + 10, 11])
            {
                rng.Style.Font.Name = "Times New Roman";
            }

            string id_contract = "";
            string id_order = "";
            int k = 0;
            for (int i = 0; i < dataGridView21.RowCount; i++)
            {
                string contract_id = "";
                string contract_date = "";
                string order_id = "";
                string order_id_date = "";
                try { contract_id = dataGridView21.Rows[i].Cells[1].Value.ToString(); } catch { contract_id = ""; }
                try { order_id = dataGridView21.Rows[i].Cells[2].Value.ToString(); } catch { order_id = ""; }
                try { order_id_date = dataGridView21.Rows[i].Cells[3].Value.ToString(); } catch { order_id_date = ""; }
                if (contract_id != id_contract)
                {
                    for (int m = 0; m < dataGridView20.RowCount; m++)
                    {
                        if (contract_id == dataGridView20.Rows[m].Cells[1].Value.ToString())
                        {
                            contract_date = dataGridView20.Rows[m].Cells[2].Value.ToString();
                            worksheet.Cells[8 + k, 8].Value = contract_id + "\n" + contract_date;
                            worksheet.Cells[8 + k, 11].Value = order_id + "\n" + order_id_date;
                            k++;
                        }
                    }
                }
                id_contract = contract_id;
                id_order = order_id;
            }


            using (ExcelRange rng = worksheet.Cells[3, 1, 20, 20])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[22, 1, rowCount, 13])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[3, 1, rowCount, 20])
            {
                rng.Style.WrapText = true;
                rng.Style.Font.Size = 8;
                for (int i = 3; i <= worksheet.Dimension.Rows + 1; i++)
                {
                    if (i == 21)
                    {
                        worksheet.Row(i).Height = 45;
                    }
                    else
                    {
                        worksheet.Row(i).Height = 23;
                    }

                }
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                excel.SaveAs(new FileInfo(saveFileDialog.FileName));
            }
        }

        public void MakeInvoiceDoc()
        {
            DateTime dateTime = DateTime.Now;
            string year = dateTime.Year.ToString();
            string month = dateTime.Month.ToString();
            string day = dateTime.Day.ToString();
            if (month.Length == 1) { month = "0" + month; }
            if (day.Length == 1) { day = "0" + day; }
            string today = day + "." + month + "." + year;

            ExcelPackage excel = new ExcelPackage();
            OfficeOpenXml.ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("sheet1");

            worksheet.Row(1).Height = 60;
            worksheet.Row(2).Height = 40;
            worksheet.Row(18).Height = 35;
            worksheet.Column(1).Width = 5;
            worksheet.Column(2).Width = 5;
            worksheet.Column(3).Width = 5;
            worksheet.Column(4).Width = 3;
            worksheet.Column(5).Width = 11;
            worksheet.Column(6).Width = 15;
            worksheet.Column(7).Width = 10;
            worksheet.Column(8).Width = 10;
            worksheet.Column(9).Width = 5;
            worksheet.Column(10).Width = 5;
            worksheet.Column(11).Width = 7;
            worksheet.Column(12).Width = 4;
            worksheet.Column(13).Width = 4;

            string year_arriving = Convert.ToString(dateTimePicker16.Value.Year);
            string month_arriving = Convert.ToString(dateTimePicker16.Value.Month);
            string day_arriving = Convert.ToString(dateTimePicker16.Value.Day);
            if (month_arriving.Length == 1) { month_arriving = "0" + month_arriving; }
            if (day_arriving.Length == 1) { day_arriving = "0" + day_arriving; }
            string arriving_date = day_arriving + "." + month_arriving + "." + year_arriving;
            worksheet.Cells[2, 1].Value = "Commercial Invoice / Коммерческий инвойс";
            worksheet.Cells[3, 1].Value = "1. Отправитель/Экспортер (Shipper/Exporter)";
            worksheet.Cells[4, 1].Value = textBoxExporterCompany.Text;
            worksheet.Cells[5, 1].Value = textBoxExporterAddress.Text;
            worksheet.Cells[6, 1].Value = textBoxExporterContact.Text;
            worksheet.Cells[7, 1].Value = "2. Получатель (Consignee)";
            worksheet.Cells[8, 1].Value = comboBox10.GetItemText(comboBox10.SelectedItem);
            worksheet.Cells[9, 1].Value = comboBox11.GetItemText(comboBox11.SelectedItem);
            worksheet.Cells[10, 1].Value = comboBox12.GetItemText(comboBox12.SelectedItem);
            worksheet.Cells[12, 1].Value = "3. Извещаемая сторона (Notify party)";
            worksheet.Cells[13, 1].Value = textBoxNotifyCompany.Text;
            worksheet.Cells[14, 1].Value = textBoxNotifyAdress.Text;
            worksheet.Cells[15, 1].Value = "4. Порт загрузки (Loading port)";
            worksheet.Cells[16, 1].Value = textBoxLoadingPort.Text;
            worksheet.Cells[15, 6].Value = "5. Порт разгрузки (Port of Discharging)";
            worksheet.Cells[16, 6].Value = comboBox13.GetItemText(comboBox13.SelectedItem);
            worksheet.Cells[17, 1].Value = "6. Пункт назначения(Final Destination)";
            worksheet.Cells[18, 1].Value = textBoxFinalDestination.Text;
            worksheet.Cells[17, 6].Value = "7. Перевозчик (Carrier)";
            worksheet.Cells[18, 6].Value = textBoxCarrier.Text;
            worksheet.Cells[19, 1].Value = "8. Номер рейса(Voyage No.)";
            worksheet.Cells[20, 1].Value = textBoxVoyageNo.Text;
            worksheet.Cells[19, 6].Value = "9. Примерная дата отгрузки (Sailing on or about)";
            worksheet.Cells[20, 6].Value = arriving_date;

            worksheet.Cells[3, 8].Value = "10. Номер и дата инвойса (No. & Date of invoice)";
            worksheet.Cells[4, 8].Value = "";
            worksheet.Cells[4, 11].Value = today;
            worksheet.Cells[7, 8].Value = "11. Примечание (Remarks)";
            worksheet.Cells[7, 11].Value = "Order ";
            worksheet.Cells[17, 8].Value = "Country of Origin: Korea";
            worksheet.Cells[18, 8].Value = "Terms of delivery : " + textBoxFinalDestination.Text;
            using (ExcelRange rng = worksheet.Cells[1, 1, 1, 13])
            {
                rng.Merge = true;
            }
            using (System.Drawing.Image logo = System.Drawing.Image.FromFile(app_dir_temp + "Logo.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("My Logo", logo);
                excelImage.SetPosition(0, 0, 5, 40);
                excelImage.SetSize(20);
            }
            using (ExcelRange rng = worksheet.Cells[2, 1, 2, 13])
            {
                rng.Merge = true;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Font.Size = 15;
                rng.Style.Font.Bold = true;
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[3, 1, 3, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[4, 1, 4, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[5, 1, 5, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[6, 1, 6, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[7, 1, 7, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[8, 1, 8, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[9, 1, 9, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[10, 1, 10, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[11, 1, 11, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[12, 1, 12, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[13, 1, 13, 7])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[14, 1, 14, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[15, 1, 15, 5])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[15, 6, 15, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[16, 1, 16, 5])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[16, 6, 16, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[17, 1, 17, 5])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[17, 6, 17, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[18, 1, 18, 5])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[18, 6, 18, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[19, 1, 19, 5])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[19, 6, 19, 7])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[20, 1, 20, 5])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[20, 6, 20, 7])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[3, 7, 20, 7])
            {
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[3, 8, 3, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[4, 8, 4, 10])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[4, 11, 4, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[5, 8, 5, 13])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[6, 8, 6, 13])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
            using (ExcelRange rng = worksheet.Cells[7, 8, 7, 10])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[7, 11, 7, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            for (int i = 8; i < 17; i++)
            {
                using (ExcelRange rng = worksheet.Cells[i, 8, i, 10])
                {
                    rng.Merge = true;
                }
                using (ExcelRange rng = worksheet.Cells[i, 11, i, 13])
                {
                    rng.Merge = true;
                }
            }
            using (ExcelRange rng = worksheet.Cells[17, 8, 17, 13])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[18, 8, 18, 13])
            {
                rng.Merge = true;
                rng.Style.Font.Bold = true;
            }
            using (ExcelRange rng = worksheet.Cells[19, 8, 19, 13])
            {
                rng.Merge = true;
            }
            using (ExcelRange rng = worksheet.Cells[20, 8, 20, 13])
            {
                rng.Merge = true;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            worksheet.Cells[21, 1].Value = "12.  Goods Items";
            worksheet.Cells[21, 1, 21, 4].Merge = true;

            worksheet.Cells[21, 5].Value = "13. Item Code";
            //worksheet.Cells[21, 5, 21, 6].Merge = true;

            worksheet.Cells[21, 6].Value = "14. Описание торвара(Description of Goods)";
            worksheet.Cells[21, 6, 21, 7].Merge = true;

            worksheet.Cells[21, 8].Value = "15. Hs Code";

            worksheet.Cells[21, 9].Value = "16. Ед. изм.(Unit)";

            worksheet.Cells[21, 10].Value = "17.Кол - во(Q'ty";

            worksheet.Cells[21, 11].Value = "18. ЦенаUnit Price(USD)";

            worksheet.Cells[21, 12].Value = "19. Сумма Amount(USD)";
            worksheet.Cells[21, 12, 21, 13].Merge = true;

            using (ExcelRange rng = worksheet.Cells[21, 1, 21, 13])
            {
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            }


            int rowCount = 22;
            int subTotalQty = 0;
            int totalQty = 0;
            double subtotalAmountPrice = 0;
            double totalAmountPrice = 0;
            int startIndex = 22;
            string idContract = "";

            for (int i = 0; i < dataGridView21.RowCount; i++)
            {
                string contract_id = "";
                try { contract_id = dataGridView21.Rows[i].Cells[1].Value.ToString(); } catch { contract_id = ""; }
                if (contract_id != idContract && subTotalQty == 0)
                {
                    idContract = contract_id;
                    startIndex = rowCount;

                    worksheet.Cells[rowCount, 1].Value = "No: " + dataGridView21.Rows[i].Cells[1].Value.ToString();

                    worksheet.Cells[rowCount, 5].Value = dataGridView21.Rows[i].Cells[4].Value.ToString();
                    worksheet.Cells[rowCount, 6].Value = dataGridView21.Rows[i].Cells[5].Value.ToString();
                    worksheet.Cells[rowCount, 6, rowCount, 7].Merge = true;
                    worksheet.Cells[rowCount, 8].Value = dataGridView21.Rows[i].Cells[6].Value.ToString();
                    worksheet.Cells[rowCount, 9].Value = "EA";
                    worksheet.Cells[rowCount, 10].Value = dataGridView21.Rows[i].Cells[7].Value.ToString();
                    worksheet.Cells[rowCount, 11].Value = dataGridView21.Rows[i].Cells[8].Value.ToString();
                    worksheet.Cells[rowCount, 12].Value = dataGridView21.Rows[i].Cells[9].Value.ToString();
                    worksheet.Cells[rowCount, 12, rowCount, 13].Merge = true;

                    subTotalQty = subTotalQty + Convert.ToInt32(dataGridView21.Rows[i].Cells[7].Value.ToString());
                    subtotalAmountPrice = subtotalAmountPrice + Math.Round(Convert.ToDouble(dataGridView21.Rows[i].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);

                    rowCount++;
                }
                else if (contract_id == idContract)
                {
                    worksheet.Cells[rowCount, 5].Value = dataGridView21.Rows[i].Cells[4].Value.ToString();
                    worksheet.Cells[rowCount, 6].Value = dataGridView21.Rows[i].Cells[5].Value.ToString();
                    worksheet.Cells[rowCount, 6, rowCount, 7].Merge = true;
                    worksheet.Cells[rowCount, 8].Value = dataGridView21.Rows[i].Cells[6].Value.ToString();
                    worksheet.Cells[rowCount, 9].Value = "EA";
                    worksheet.Cells[rowCount, 10].Value = dataGridView21.Rows[i].Cells[7].Value.ToString();
                    worksheet.Cells[rowCount, 11].Value = dataGridView21.Rows[i].Cells[8].Value.ToString();
                    worksheet.Cells[rowCount, 12].Value = dataGridView21.Rows[i].Cells[9].Value.ToString();
                    worksheet.Cells[rowCount, 12, rowCount, 13].Merge = true;

                    subTotalQty = subTotalQty + Convert.ToInt32(dataGridView21.Rows[i].Cells[7].Value.ToString());
                    subtotalAmountPrice = subtotalAmountPrice + Math.Round(Convert.ToDouble(dataGridView21.Rows[i].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);

                    rowCount++;
                }
                else if (contract_id != idContract && subTotalQty > 0)
                {
                    worksheet.Cells[startIndex, 1, rowCount, 4].Merge = true;
                    worksheet.Cells[rowCount, 5, rowCount, 9].Merge = true;
                    worksheet.Cells[rowCount, 5].Value = "SUBTOTAL FOR #" + contract_id.ToString();
                    worksheet.Cells[rowCount, 10].Value = subTotalQty + " EA";
                    worksheet.Cells[rowCount, 12].Value = "$" + subtotalAmountPrice;
                    worksheet.Cells[rowCount, 12, rowCount, 13].Merge = true;
                    using (ExcelRange rng = worksheet.Cells[rowCount, 5, rowCount, 13])
                    {
                        rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                    totalQty = totalQty + subTotalQty;
                    totalAmountPrice = totalAmountPrice + subtotalAmountPrice;
                    subtotalAmountPrice = 0;
                    subTotalQty = 0;

                    rowCount++;
                    startIndex = rowCount;

                    worksheet.Cells[rowCount, 1].Value = "No: " + dataGridView21.Rows[i].Cells[1].Value.ToString();

                    worksheet.Cells[rowCount, 5].Value = dataGridView21.Rows[i].Cells[4].Value.ToString();
                    worksheet.Cells[rowCount, 6].Value = dataGridView21.Rows[i].Cells[5].Value.ToString();
                    worksheet.Cells[rowCount, 6, rowCount, 7].Merge = true;
                    worksheet.Cells[rowCount, 8].Value = dataGridView21.Rows[i].Cells[6].Value.ToString();
                    worksheet.Cells[rowCount, 9].Value = "EA";
                    worksheet.Cells[rowCount, 10].Value = dataGridView21.Rows[i].Cells[7].Value.ToString();
                    worksheet.Cells[rowCount, 11].Value = dataGridView21.Rows[i].Cells[8].Value.ToString();
                    worksheet.Cells[rowCount, 12].Value = dataGridView21.Rows[i].Cells[9].Value.ToString();
                    worksheet.Cells[rowCount, 12, rowCount, 13].Merge = true;

                    subTotalQty = subTotalQty + Convert.ToInt32(dataGridView21.Rows[i].Cells[7].Value.ToString());
                    subtotalAmountPrice = subtotalAmountPrice + Math.Round(Convert.ToDouble(dataGridView21.Rows[i].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);

                    rowCount++;
                    idContract = contract_id;
                }
            }

            worksheet.Cells[startIndex, 1, rowCount, 4].Merge = true;
            worksheet.Cells[rowCount, 5, rowCount, 9].Merge = true;
            worksheet.Cells[rowCount, 5].Value = "SUBTOTAL FOR #" + idContract.ToString();
            worksheet.Cells[rowCount, 10].Value = subTotalQty + " EA";
            worksheet.Cells[rowCount, 12].Value = "$" + subtotalAmountPrice;
            worksheet.Cells[rowCount, 12, rowCount, 13].Merge = true;
            using (ExcelRange rng = worksheet.Cells[rowCount, 5, rowCount, 13])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }
            totalQty = totalQty + subTotalQty;
            totalAmountPrice = totalAmountPrice + subtotalAmountPrice;
            subtotalAmountPrice = 0;
            subTotalQty = 0;
            rowCount++;

            worksheet.Cells[rowCount, 1, rowCount, 9].Merge = true;
            worksheet.Cells[rowCount, 12, rowCount, 13].Merge = true;
            worksheet.Cells[rowCount, 1].Value = "TOTAL: ";
            worksheet.Cells[rowCount, 10].Value = totalQty + " EA";
            worksheet.Cells[rowCount, 12].Value = "$" + totalAmountPrice;
            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 13])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
            }
            using (ExcelRange rng = worksheet.Cells[rowCount, 1, rowCount, 13])
            {
                rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
            }
            using (System.Drawing.Image sign = System.Drawing.Image.FromFile(app_dir_temp + "Sign.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Sign", sign);
                excelImage.SetPosition(rowCount + 3, 20, 6, 0);
                excelImage.SetSize(18);
            }
            using (System.Drawing.Image stamp = System.Drawing.Image.FromFile(app_dir_temp + "Stamp.png"))
            {
                var excelImage = worksheet.Drawings.AddPicture("Stamp", stamp);
                excelImage.SetPosition(rowCount + 3, 0, 8, 20);
                excelImage.SetSize(35);
            }

            worksheet.Cells[rowCount + 6, 6, rowCount + 6, 11].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[rowCount + 7, 6].Value = " Authorized and signed by Anna Li";
            worksheet.Cells[rowCount + 7, 6, rowCount + 7, 9].Merge = true;
            worksheet.Cells[rowCount + 8, 6].Value = " by Power of Attorney on behalf of ";
            worksheet.Cells[rowCount + 8, 6, rowCount + 8, 9].Merge = true;
            worksheet.Cells[rowCount + 9, 6].Value = " Drone Zone General Manager  ";
            worksheet.Cells[rowCount + 9, 6, rowCount + 9, 9].Merge = true;
            worksheet.Cells[rowCount + 10, 6].Value = " dronezone.anna@gmail.com  ";
            worksheet.Cells[rowCount + 10, 6].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 10, 6, rowCount + 10, 9].Merge = true;

            worksheet.Cells[rowCount + 9, 1].Value = " Customer care  ";
            worksheet.Cells[rowCount + 9, 1, rowCount + 9, 4].Merge = true;
            worksheet.Cells[rowCount + 10, 1].Value = " dronezone.sk@gmail.com  ";
            worksheet.Cells[rowCount + 10, 1].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[rowCount + 10, 1, rowCount + 10, 5].Merge = true;

            worksheet.Cells[rowCount + 7, 10].Value = today;
            worksheet.Cells[rowCount + 7, 10, rowCount + 7, 11].Merge = true;
            using (ExcelRange rng = worksheet.Cells[rowCount + 7, 1, rowCount + 10, 11])
            {
                rng.Style.Font.Name = "Times New Roman";
            }


            string id_contract = "";
            string id_order = "";
            int k = 0;
            for (int i = 0; i < dataGridView21.RowCount; i++)
            {
                string contract_id = "";
                string contract_date = "";
                string order_id = "";
                string order_id_date = "";
                try { contract_id = dataGridView21.Rows[i].Cells[1].Value.ToString(); } catch { contract_id = ""; }
                try { order_id = dataGridView21.Rows[i].Cells[2].Value.ToString(); } catch { order_id = ""; }
                try { order_id_date = dataGridView21.Rows[i].Cells[3].Value.ToString(); } catch { order_id_date = ""; }
                if (contract_id != id_contract)
                {
                    for (int m = 0; m < dataGridView20.RowCount; m++)
                    {
                        if (contract_id == dataGridView20.Rows[m].Cells[1].Value.ToString())
                        {
                            contract_date = dataGridView20.Rows[m].Cells[2].Value.ToString();
                            worksheet.Cells[8 + k, 8].Value = contract_id + "\n" + contract_date;
                            worksheet.Cells[8 + k, 11].Value = order_id + "\n" + order_id_date;
                            k++;
                        }
                    }
                }
                id_contract = contract_id;
                id_order = order_id;
            }


            using (ExcelRange rng = worksheet.Cells[3, 1, 20, 20])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[22, 1, rowCount, 13])
            {
                rng.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                rng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rng.Style.Font.Name = "Times New Roman";
            }
            using (ExcelRange rng = worksheet.Cells[3, 1, rowCount, 20])
            {
                rng.Style.WrapText = true;
                rng.Style.Font.Size = 8;
                for (int i = 3; i <= worksheet.Dimension.Rows + 1; i++)
                {
                    if (i == 21)
                    {
                        worksheet.Row(i).Height = 45;
                    }
                    else
                    {
                        worksheet.Row(i).Height = 23;
                    }

                }
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                excel.SaveAs(new FileInfo(saveFileDialog.FileName));
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            if (comboBox9.SelectedIndex == 0)
            {
                GetContractListToGenerate(textBox34.Text, 0);
            }
            if (comboBox9.SelectedIndex == 1)
            {
                string year = Convert.ToString(dateTimePicker17.Value.Year);
                string month = Convert.ToString(dateTimePicker17.Value.Month);
                string day = Convert.ToString(dateTimePicker17.Value.Day);
                if (month.Length == 1) { month = "0" + month; }
                if (day.Length == 1) { day = "0" + day; }
                string contract_date = day + "." + month + "." + year;
                GetContractListToGenerate(contract_date, 1);
            }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            dataGridView21.RowCount = 0;
        }

        private void button46_Click(object sender, EventArgs e)
        {
            int selected = 0;
            selected = dataGridView20.CurrentCell.RowIndex;
            string id_contract = dataGridView20.Rows[selected].Cells[1].Value.ToString();
            bool check = true;
            for (int i = 0; i < dataGridView21.RowCount; i++)
            {
                string contract = "";
                try { contract = dataGridView21.Rows[i].Cells[1].Value.ToString(); } catch { contract = ""; }
                if (contract == id_contract)
                {
                    check = false;
                }
            }
            if (check == true)
            {
                getCollectedContractItems(id_contract);
            }
            setCompanyList();
        }
        public class companyList
        {
            public string company_name { get; set; }
        }
        public void setCompanyList()
        {
            comboBox10.Items.Clear();

            List<companyList> list = new List<companyList>();
            string sql = "SELECT DISTINCT company FROM company_details";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            companyList company = new companyList();
                            company.company_name = reader.GetValue(0).ToString();
                            list.Add(company);
                        }
                    }
                }
                connection.Close();
            }
            for(int i=0; i<list.Count; i++)
            {
                comboBox10.Items.Add(list[i].company_name);
            }
        }
        private void button44_Click(object sender, EventArgs e)
        {
            MakePLDoc();
        }

        private void button45_Click(object sender, EventArgs e)
        {
            MakeInvoiceDoc();
        }

        private void label82_Click(object sender, EventArgs e)
        {
            textBox32.Clear();
        }

        private void label107_Click(object sender, EventArgs e)
        {
            textBox36.Clear();
        }

        private void dataGridView9_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            // axad 22.07 start
            for (int i = 0; i < dataGridView9.RowCount - 1; i++)
            {
                dataGridView9.Rows[i].Cells[1].Value = i + 1;
            }
            // axad 22.07 end
            int indexOfTotalRow = dataGridView9.RowCount - 1;
            dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = 0;

            dataGridView9.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView9.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        private void dataGridView12_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView12.CurrentCell.RowIndex == dataGridView12.RowCount - 1)
            {
                dataGridView12.AllowUserToDeleteRows = false;
            }
            else
            {
                dataGridView12.AllowUserToDeleteRows = true;
            }
        }

        private void dataGridView12_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            // axad 22.07 start
            for(int i = 0; i<dataGridView12.RowCount - 1; i++)
            {
                dataGridView12.Rows[i].Cells[1].Value = i + 1;
            }
            // axad 22.07 end
            int indexOfTotalRow = dataGridView12.RowCount - 1;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;

            dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        private void dataGridView9_DoubleClick(object sender, EventArgs e)
        {
            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView9.RowCount == 0) {
                dataGridView9.RowCount = 2;
            } else {
                dataGridView9.RowCount = dataGridView9.RowCount + 1;
            }
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[1].Value = dataGridView9.RowCount - 1;
            if (dataGridView9.RowCount % 2 == 0)
            {
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[0].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[1].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[2].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[3].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[4].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[5].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[6].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[7].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[9].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[10].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[11].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[12].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[13].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[14].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[15].Style.BackColor = Color.SandyBrown; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[16].Style.BackColor = Color.SandyBrown; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[17].Style.BackColor = Color.SkyBlue;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[18].Style.BackColor = Color.SkyBlue;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[19].Style.BackColor = Color.SandyBrown;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[20].Style.BackColor = Color.SkyBlue;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[21].Style.BackColor = Color.SandyBrown; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[22].Style.BackColor = Color.SkyBlue; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[23].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[24].Style.BackColor = cl_even;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[25].Style.BackColor = cl_even;
            }
            else
            {
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[0].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[1].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[2].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[3].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[4].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[5].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[6].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[7].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[9].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[10].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[11].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[12].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[13].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[14].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[15].Style.BackColor = Color.SandyBrown; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[16].Style.BackColor = Color.SandyBrown; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[17].Style.BackColor = Color.LightSkyBlue;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[18].Style.BackColor = Color.LightSkyBlue;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[19].Style.BackColor = Color.SandyBrown;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[20].Style.BackColor = Color.LightSkyBlue;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[21].Style.BackColor = Color.SandyBrown; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[23].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[24].Style.BackColor = cl_odd;
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[25].Style.BackColor = cl_odd;
                // SeaGreen MediumSeaGreen DarkSeaGreen
                // AliceBlue CornflowerBlue AliceBlue
            }
            //
            for (int n = 7; n < 23; n++)
            {
                dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[n].Value = 0;
            }
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[2].Value = "";
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[3].Value = "";
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[5].Value = 0;
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[6].Value = 0;
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[8].Value = 71;
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[23].Value = "2-3 week";
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[24].Value = "2-3 days";
            dataGridView9.Rows[dataGridView9.RowCount - 2].Cells[25].Value = "3-4 week";
            //
            int indexOfTotalRow = dataGridView9.RowCount - 1;
            dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = 0;

            dataGridView9.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView9.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView9.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView9.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView9.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView9.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView9.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView9.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView9.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView9.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (dataGridView12.RowCount == 0){
                dataGridView12.RowCount = 2;
            }else{
                dataGridView12.RowCount = dataGridView12.RowCount + 1;
            }
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[1].Value = dataGridView12.RowCount - 1;

            if (dataGridView12.RowCount % 2 == 0)
            {
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[0].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[1].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[2].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[3].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[4].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[5].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[6].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[7].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[9].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[10].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[11].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[12].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[13].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[14].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[15].Style.BackColor = Color.SandyBrown; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[16].Style.BackColor = Color.SandyBrown; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[17].Style.BackColor = Color.SkyBlue;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[18].Style.BackColor = Color.SkyBlue;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[19].Style.BackColor = Color.SandyBrown;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[20].Style.BackColor = Color.SkyBlue;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[21].Style.BackColor = Color.SandyBrown; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[22].Style.BackColor = Color.SkyBlue; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[23].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[24].Style.BackColor = cl_even;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[25].Style.BackColor = cl_even;
            }
            else
            {
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[0].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[1].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[2].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[3].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[4].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[5].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[6].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[7].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[8].Style.BackColor = Color.MediumSeaGreen; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[9].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[10].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[11].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[12].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[13].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[14].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[15].Style.BackColor = Color.SandyBrown; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[16].Style.BackColor = Color.SandyBrown; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[17].Style.BackColor = Color.LightSkyBlue;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[18].Style.BackColor = Color.LightSkyBlue;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[19].Style.BackColor = Color.SandyBrown;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[20].Style.BackColor = Color.LightSkyBlue;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[21].Style.BackColor = Color.SandyBrown; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[22].Style.BackColor = Color.LightSkyBlue; //
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[23].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[24].Style.BackColor = cl_odd;
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[25].Style.BackColor = cl_odd;
                // SeaGreen MediumSeaGreen DarkSeaGreen
                // AliceBlue CornflowerBlue AliceBlue
            }
            //
            for (int n = 7; n < 23; n++)
            {
                dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[n].Value = 0;
            }
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[2].Value = "";
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[3].Value = "";
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[5].Value = 0;
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[6].Value = 0;
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[8].Value = 71;
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[23].Value = "2-3 week";
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[24].Value = "2-3 days";
            dataGridView12.Rows[dataGridView12.RowCount - 2].Cells[25].Value = "3-4 week";
            //
            int indexOfTotalRow = dataGridView12.RowCount - 1;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = 0;
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = 0;


            dataGridView12.Rows[indexOfTotalRow].Cells[0].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[1].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[2].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[3].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[4].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[5].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[7].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[10].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[12].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[14].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Style.BackColor = Color.DimGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[16].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[18].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Style.BackColor = Color.DimGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[21].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[22].Style.BackColor = Color.LightSlateGray; //
            dataGridView12.Rows[indexOfTotalRow].Cells[23].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[24].Style.BackColor = Color.LightSlateGray;
            dataGridView12.Rows[indexOfTotalRow].Cells[25].Style.BackColor = Color.LightSlateGray;

            double counter = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView12.Rows[indexOfTotalRow].Cells[3].Value = "TOTAL: ";
                dataGridView12.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[7].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[7].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[10].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[10].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[12].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[12].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[14].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[14].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[16].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[16].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[18].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[18].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[21].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[21].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                dataGridView12.Rows[indexOfTotalRow].Cells[22].Value = Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) + Math.Round(Convert.ToDouble(dataGridView12.Rows[k].Cells[22].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                counter++;
            }
            dataGridView12.Rows[indexOfTotalRow].Cells[6].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[6].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[8].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[8].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[9].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[9].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[11].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[11].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[13].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[13].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[17].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[17].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[15].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[15].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[19].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[19].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
            dataGridView12.Rows[indexOfTotalRow].Cells[20].Value = Math.Round(Math.Round(Convert.ToDouble(dataGridView12.Rows[indexOfTotalRow].Cells[20].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven) / counter, 2, MidpointRounding.ToEven);
        }

        private void button48_Click(object sender, EventArgs e)
        {
            Random rnd = new Random();
            int val = rnd.Next(9999, 99999);
            File.Copy(app_dir_temp + "tpl.xlsx", app_dir_temp + "templ" + val + ".xlsx");
            System.Diagnostics.Process.Start(app_dir_temp + "templ" + val + ".xlsx");

            dataGridView22.RowCount = dataGridView22.RowCount + 1;
            dataGridView22.Rows[dataGridView22.RowCount - 1].Cells[0].Value = dataGridView22.RowCount;
            dataGridView22.Rows[dataGridView22.RowCount - 1].Cells[1].Value = "Шаблон " + dataGridView22.RowCount;
            dataGridView22.Rows[dataGridView22.RowCount - 1].Cells[2].Value = app_dir_temp + "templ" + val + ".xlsx";

            dataGridView23.RowCount = dataGridView23.RowCount + 1;
            dataGridView23.Rows[dataGridView23.RowCount - 1].Cells[0].Value = dataGridView23.RowCount;
            dataGridView23.Rows[dataGridView23.RowCount - 1].Cells[1].Value = "Шаблон " + dataGridView23.RowCount;
            dataGridView23.Rows[dataGridView23.RowCount - 1].Cells[2].Value = app_dir_temp + "templ" + val + ".xlsx";
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (dataGridView22.RowCount != 0)
            {
                int selected = dataGridView22.CurrentCell.RowIndex;
                string file = dataGridView22.Rows[selected].Cells[2].Value.ToString();
                LoadCPGrid(file);
            }
        }

        private void открытьШаблонВEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView22.RowCount != 0)
            {
                int selected = dataGridView22.CurrentCell.RowIndex;
                string file = dataGridView22.Rows[selected].Cells[2].Value.ToString();
                System.Diagnostics.Process.Start(file);
            }
        }

        private void LoadRootFolder()
        {
            try {
                StreamReader fs = new StreamReader(app_dir_temp + "root.txt");
                string s = "";
                while (s != null)
                {
                    s = fs.ReadLine();
                    break;
                }
                fs.Close();
                contextMenuStrip7.Items[1].Text = contextMenuStrip7.Items[1].Text + s;
            } catch { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ShowPanels("Ком");
            InitGridView12();
            InitGridView9();
            InitGrid10();
            GetCPList("", 0);
            ClearCash();
            LoadRootFolder();
            FindCompanies();
        }

        private void button50_Click(object sender, EventArgs e)
        {
            Frm6.Show();
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            findAddresses();
        }
        public class addressList
        {
            public string address { get; set; }
        }
        public void findAddresses()
        {
            comboBox11.Items.Clear();

            string companyName = comboBox10.GetItemText(comboBox10.SelectedItem);

            List<addressList> address_list = new List<addressList>();
            string sql = "SELECT address FROM company_details WHERE company='" + companyName + "'";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            addressList address = new addressList();
                            address.address = reader.GetValue(0).ToString();
                            address_list.Add(address);
                        }
                    }
                }
                connection.Close();
            }

            for(int i=0; i<address_list.Count; i++)
            {
                comboBox11.Items.Add(address_list[i].address);
            }
            comboBox11.SelectedIndex = -1;
            comboBox12.SelectedIndex = -1;
            comboBox13.SelectedIndex = -1;
            textBoxNotifyCompany.Text = "";
            textBoxNotifyAdress.Text = "";
            textBoxNotifyContact.Text = "";
            textBoxFinalDestination.Text = "";
            textBoxConditions.Text = "";
        }
        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            findContact();
        }
        public class contactList
        {
            public string contact { get; set; }
        }
        public void findContact()
        {
            comboBox12.Items.Clear();

            string companyName = comboBox10.GetItemText(comboBox10.SelectedItem);
            string companyAddress = comboBox11.GetItemText(comboBox11.SelectedItem);

            List<contactList> contact_list = new List<contactList>();
            string sql = "SELECT contact FROM company_details WHERE company='" + companyName + "' and address='"+companyAddress+"'";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            contactList contact = new contactList();
                            contact.contact = reader.GetValue(0).ToString();
                            contact_list.Add(contact);
                        }
                    }
                }
                connection.Close();
            }

            for (int i = 0; i < contact_list.Count; i++)
            {
                comboBox12.Items.Add(contact_list[i].contact);
            }
            comboBox12.SelectedIndex = -1;
            comboBox13.SelectedIndex = -1;
            textBoxNotifyCompany.Text = "";
            textBoxNotifyAdress.Text = "";
            textBoxNotifyContact.Text = "";
            textBoxFinalDestination.Text = "";
            textBoxConditions.Text = "";
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxNotifyCompany.Text = comboBox10.GetItemText(comboBox10.SelectedItem);
            textBoxNotifyAdress.Text = comboBox11.GetItemText(comboBox11.SelectedItem);
            textBoxNotifyContact.Text = comboBox12.GetItemText(comboBox12.SelectedItem);
            findPorts();
        }
        public class portsList
        {
            public string port { get; set; }
        }
        public void findPorts()
        {
            comboBox13.Items.Clear();

            string companyName = comboBox10.GetItemText(comboBox10.SelectedItem);
            string companyAddress = comboBox11.GetItemText(comboBox11.SelectedItem);
            string companyContact = comboBox12.GetItemText(comboBox12.SelectedItem);

            List<portsList> ports_list = new List<portsList>();
            string sql = "SELECT port_name FROM company_details WHERE company='" + companyName + "' and address='" + companyAddress + "' and contact='"+companyContact+"'";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            portsList port = new portsList();
                            port.port = reader.GetValue(0).ToString();
                            ports_list.Add(port);
                        }
                    }
                }
                connection.Close();
            }

            for (int i = 0; i < ports_list.Count; i++)
            {
                comboBox13.Items.Add(ports_list[i].port);
            }
            comboBox13.SelectedIndex = -1;
            textBoxFinalDestination.Text = "";
            textBoxConditions.Text = "";
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            findFDestination();
        }
        public void findFDestination()
        {
            string companyName = comboBox10.GetItemText(comboBox10.SelectedItem);
            string companyAddress = comboBox11.GetItemText(comboBox11.SelectedItem);
            string companyContact = comboBox12.GetItemText(comboBox12.SelectedItem);
            string companyPort = comboBox13.GetItemText(comboBox13.SelectedItem);
            string finalDestination = "";

            string sql = "SELECT final_destination FROM company_details WHERE company='" + companyName + "' and address='" + companyAddress + "' and contact='" + companyContact + "' and port_name='"+companyPort+"'";
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            finalDestination = reader.GetValue(0).ToString(); 
                        }
                    }
                }
                connection.Close();
            }
            textBoxFinalDestination.Text = finalDestination;
            textBoxConditions.Text = finalDestination;
        }

        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView9.CurrentCell.RowIndex == dataGridView9.RowCount - 1)
            {
                dataGridView9.AllowUserToDeleteRows = false;
            }
            else
            {
                dataGridView9.AllowUserToDeleteRows = true;
            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (dataGridView23.RowCount != 0)
            {
                int selected = dataGridView23.CurrentCell.RowIndex;
                string file = dataGridView23.Rows[selected].Cells[2].Value.ToString();
                LoadCPGrid2(file);
            }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (dataGridView23.RowCount != 0)
            {
                int selected = dataGridView23.CurrentCell.RowIndex;
                string file = dataGridView23.Rows[selected].Cells[2].Value.ToString();
                System.Diagnostics.Process.Start(file);
            }
        }

        private void button52_Click(object sender, EventArgs e)
        {
            Random rnd = new Random();
            int val = rnd.Next(9999, 99999);
            File.Copy(app_dir_temp + "tpl.xlsx", app_dir_temp + "templ" + val + ".xlsx");
            System.Diagnostics.Process.Start(app_dir_temp + "templ" + val + ".xlsx");

            dataGridView22.RowCount = dataGridView22.RowCount + 1;
            dataGridView22.Rows[dataGridView22.RowCount - 1].Cells[0].Value = dataGridView22.RowCount;
            dataGridView22.Rows[dataGridView22.RowCount - 1].Cells[1].Value = "Шаблон " + dataGridView22.RowCount;
            dataGridView22.Rows[dataGridView22.RowCount - 1].Cells[2].Value = app_dir_temp + "templ" + val + ".xlsx";

            dataGridView23.RowCount = dataGridView23.RowCount + 1;
            dataGridView23.Rows[dataGridView23.RowCount - 1].Cells[0].Value = dataGridView23.RowCount;
            dataGridView23.Rows[dataGridView23.RowCount - 1].Cells[1].Value = "Шаблон " + dataGridView23.RowCount;
            dataGridView23.Rows[dataGridView23.RowCount - 1].Cells[2].Value = app_dir_temp + "templ" + val + ".xlsx";
        }

        private void cRMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Frm7.Show();
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
            comboBox16.Items.Clear();
            comboBox17.Items.Clear();
            for (int i = 0; i < dest_list.Count; i++)
            {
                if (dest_list[i].type == "AIR") {
                    comboBox17.Items.Add(dest_list[i].dest);
                } else {
                    comboBox16.Items.Add(dest_list[i].dest);
                }
            }
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            int si = comboBox14.SelectedIndex;
            comboBox15.SelectedIndex = si;
            int c_id = Convert.ToInt32(comboBox15.Text);
            FindCompanyDest(c_id);
            comboBox16.SelectedIndex = 0;
            comboBox17.SelectedIndex = 0;
        }

        private void panel63_Click(object sender, EventArgs e)
        {
            // Panel Com Proposal
            panel_tab2_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab3_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab1_button.BackgroundImage = Properties.Resources.tab_active;
            panel_cp_tab2.Visible = false;
            panel_cp_tab2.Dock = DockStyle.None;
            panel_cp_tab3.Visible = false;
            panel_cp_tab3.Dock = DockStyle.None;
            panel_cp_tab1.Visible = true;
            panel_cp_tab1.Dock = DockStyle.Fill;
        }

        private void panel64_Click(object sender, EventArgs e)
        {
            // Panel Com Proposal
            panel_tab1_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab3_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab2_button.BackgroundImage = Properties.Resources.tab_active;
            panel_cp_tab1.Visible = false;
            panel_cp_tab1.Dock = DockStyle.None;
            panel_cp_tab3.Visible = false;
            panel_cp_tab3.Dock = DockStyle.None;
            panel_cp_tab2.Visible = true;
            panel_cp_tab2.Dock = DockStyle.Fill;
        }

        private void panel66_Click(object sender, EventArgs e)
        {
            // Panel Com Proposal
            panel_tab1_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab2_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab3_button.BackgroundImage = Properties.Resources.tab_active;
            panel_cp_tab1.Visible = false;
            panel_cp_tab1.Dock = DockStyle.None;
            panel_cp_tab2.Visible = false;
            panel_cp_tab2.Dock = DockStyle.None;
            panel_cp_tab3.Visible = true;
            panel_cp_tab3.Dock = DockStyle.Fill;
        }

        private void panel63_Click_1(object sender, EventArgs e)
        {
            if (panel_main_menu.Width == 350)
            {
                Size sz = new Size();
                sz.Width = 65;
                panel_main_menu.Width = sz.Width;
            }
            else
            {
                Size sz = new Size();
                sz.Width = 350;
                panel_main_menu.Width = sz.Width;
            }
        }

        private void label109_Click(object sender, EventArgs e)
        {
            // Panel Com Proposal
            panel_tab2_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab3_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab1_button.BackgroundImage = Properties.Resources.tab_active;
            panel_cp_tab2.Visible = false;
            panel_cp_tab2.Dock = DockStyle.None;
            panel_cp_tab3.Visible = false;
            panel_cp_tab3.Dock = DockStyle.None;
            panel_cp_tab1.Visible = true;
            panel_cp_tab1.Dock = DockStyle.Fill;
        }

        private void label110_Click(object sender, EventArgs e)
        {
            // Panel Com Proposal
            panel_tab1_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab3_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab2_button.BackgroundImage = Properties.Resources.tab_active;
            panel_cp_tab1.Visible = false;
            panel_cp_tab1.Dock = DockStyle.None;
            panel_cp_tab3.Visible = false;
            panel_cp_tab3.Dock = DockStyle.None;
            panel_cp_tab2.Visible = true;
            panel_cp_tab2.Dock = DockStyle.Fill;
        }

        private void label111_Click(object sender, EventArgs e)
        {
            // Panel Com Proposal
            panel_tab1_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab2_button.BackgroundImage = Properties.Resources.tab_inactive;
            panel_tab3_button.BackgroundImage = Properties.Resources.tab_active;
            panel_cp_tab1.Visible = false;
            panel_cp_tab1.Dock = DockStyle.None;
            panel_cp_tab2.Visible = false;
            panel_cp_tab2.Dock = DockStyle.None;
            panel_cp_tab3.Visible = true;
            panel_cp_tab3.Dock = DockStyle.Fill;
            GetContractList("", 0);
        }

        private void panel63_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel72_MouseHover(object sender, EventArgs e)
        {

        }

        private void panel_menu_home_MouseLeave(object sender, EventArgs e)
        {
        }

        // Menu Home
        private void panel_menu_home_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_home.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_home.ForeColor = Color.Black;
            //
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
        }

        private void label_menu_home_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_home.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_home.ForeColor = Color.Black;
            //
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
        }

        // Menu Ware House
        private void panel_menu_wh_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_wh.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Стелажи");
            MakeShelfsStructure(0);
            panel16.Width = 15;
        }

        private void label_menu_wh_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_wh.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Стелажи");
            MakeShelfsStructure(0);
            panel16.Width = 15;
        }

        // Menu Production Order
        private void panel_menu_po_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_po.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_po.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Заказ");
            InitGridPO();
        }

        private void label_menu_po_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_po.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_po.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Заказ");
            InitGridPO();
        }

        // Menu Income
        private void panel_menu_in_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_in.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_in.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Приход");
            InitGridIn();
        }

        private void label_menu_in_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_in.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_in.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Приход");
            InitGridIn();
        }

        // Menu Com Proposal
        private void panel_menu_cp_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_cp.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Ком");
        }

        private void label_menu_cp_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_cp.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            label_menu_basket.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Ком");
        }

        // Menu Basket
        private void panel_menu_basket_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_basket.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Корзина");
            if (dataGridView15.RowCount == 0)
            {
                InitGrid();
            }
        }

        private void label_menu_basket_Click(object sender, EventArgs e)
        {
            // 
            panel_menu_basket.BackgroundImage = Properties.Resources.menu_hover;
            label_menu_basket.ForeColor = Color.Black;
            //
            panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            label_menu_home.ForeColor = cl_odd;
            panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            label_menu_wh.ForeColor = cl_odd;
            panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            label_menu_po.ForeColor = cl_odd;
            panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            label_menu_in.ForeColor = cl_odd;
            panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            label_menu_cp.ForeColor = cl_odd;
            panel_menu_crm.BackgroundImage = Properties.Resources.menu_back;
            label_menu_crm.ForeColor = cl_odd;
            //
            ShowPanels("Корзина");
            if (dataGridView15.RowCount == 0)
            {
                InitGrid();
            }
        }

        // Menu CRM
        private void panel_menu_crm_Click(object sender, EventArgs e)
        {
            // 
            //panel_menu_crm.BackgroundImage = Properties.Resources.menu_hover;
            //label_menu_crm.ForeColor = Color.Black;
            ////
            //panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_home.ForeColor = cl_odd;
            //panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_wh.ForeColor = cl_odd;
            //panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_po.ForeColor = cl_odd;
            //panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_in.ForeColor = cl_odd;
            //panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_cp.ForeColor = cl_odd;
            //panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_basket.ForeColor = cl_odd;
            //
            Frm7.Show();
            Frm7.TopMost = false;
        }

        private void label_menu_crm_Click(object sender, EventArgs e)
        {
            // 
            //panel_menu_crm.BackgroundImage = Properties.Resources.menu_hover;
            //label_menu_crm.ForeColor = Color.Black;
            ////
            //panel_menu_home.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_home.ForeColor = cl_odd;
            //panel_menu_wh.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_wh.ForeColor = cl_odd;
            //panel_menu_po.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_po.ForeColor = cl_odd;
            //panel_menu_in.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_in.ForeColor = cl_odd;
            //panel_menu_cp.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_cp.ForeColor = cl_odd;
            //panel_menu_basket.BackgroundImage = Properties.Resources.menu_back;
            //label_menu_basket.ForeColor = cl_odd;
            Frm7.Show();
            Frm7.TopMost = false;
        }

        // Basket panel tabs
        private void panel69_Click(object sender, EventArgs e)
        {
            //
            panel_basket_tab1.Visible = true;
            panel_basket_tab1.Dock = DockStyle.Fill;
            //
            panel_basket_tab2.Visible = false;
            panel_basket_tab2.Dock = DockStyle.None;
            //
            panel_basket_tab1_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_basket_tab2_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label114_Click(object sender, EventArgs e)
        {
            //
            panel_basket_tab1.Visible = true;
            panel_basket_tab1.Dock = DockStyle.Fill;
            //
            panel_basket_tab2.Visible = false;
            panel_basket_tab2.Dock = DockStyle.None;
            //
            panel_basket_tab1_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_basket_tab2_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void panel68_Click(object sender, EventArgs e)
        {
            //
            panel_basket_tab2.Visible = true;
            panel_basket_tab2.Dock = DockStyle.Fill;
            //
            panel_basket_tab1.Visible = false;
            panel_basket_tab1.Dock = DockStyle.None;
            //
            panel_basket_tab2_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_basket_tab1_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label113_Click(object sender, EventArgs e)
        {
            //
            panel_basket_tab2.Visible = true;
            panel_basket_tab2.Dock = DockStyle.Fill;
            //
            panel_basket_tab1.Visible = false;
            panel_basket_tab1.Dock = DockStyle.None;
            //
            panel_basket_tab2_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_basket_tab1_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        // Income panel tabs
        private void panel_tab1_btn_Click(object sender, EventArgs e)
        {
            //
            panel_in_tab1.Visible = true;
            panel_in_tab1.Dock = DockStyle.Fill;
            //
            panel_in_tab2.Visible = false;
            panel_in_tab2.Dock = DockStyle.None;
            //
            panel_in_tab1_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_in_tab2_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label115_Click(object sender, EventArgs e)
        {
            //
            panel_in_tab1.Visible = true;
            panel_in_tab1.Dock = DockStyle.Fill;
            //
            panel_in_tab2.Visible = false;
            panel_in_tab2.Dock = DockStyle.None;
            //
            panel_in_tab1_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_in_tab2_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void panel_tab2_btn_Click(object sender, EventArgs e)
        {
            //
            panel_in_tab2.Visible = true;
            panel_in_tab2.Dock = DockStyle.Fill;
            //
            panel_in_tab1.Visible = false;
            panel_in_tab1.Dock = DockStyle.None;
            //
            panel_in_tab2_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_in_tab1_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label112_Click(object sender, EventArgs e)
        {
            //
            panel_in_tab2.Visible = true;
            panel_in_tab2.Dock = DockStyle.Fill;
            //
            panel_in_tab1.Visible = false;
            panel_in_tab1.Dock = DockStyle.None;
            //
            panel_in_tab2_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_in_tab1_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label118_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
        }

        private void label116_Click(object sender, EventArgs e)
        {
            textBox7.Clear();
        }

        private void label117_Click(object sender, EventArgs e)
        {
            textBox9.Clear();
        }

        private void button53_Click(object sender, EventArgs e)
        {

        }

        private void timer_panel_cover_Tick(object sender, EventArgs e)
        {
            timer_panel_cover.Enabled = false;
            Point np = new Point();
            np.X = 10000;
            np.Y = 0;
            //
            panel_cover.Location = np;
        }

        // PO panel tabs
        private void panel_po_tab1_btn_Click(object sender, EventArgs e)
        {
            //
            panel_po_tab1.Visible = true;
            panel_po_tab1.Dock = DockStyle.Fill;
            //
            panel_po_tab2.Visible = false;
            panel_po_tab2.Dock = DockStyle.None;
            //
            panel_po_tab1_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_po_tab2_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label120_Click(object sender, EventArgs e)
        {
            //
            panel_po_tab1.Visible = true;
            panel_po_tab1.Dock = DockStyle.Fill;
            //
            panel_po_tab2.Visible = false;
            panel_po_tab2.Dock = DockStyle.None;
            //
            panel_po_tab1_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_po_tab2_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void panel_po_tab2_btn_Click(object sender, EventArgs e)
        {
            //
            panel_po_tab2.Visible = true;
            panel_po_tab2.Dock = DockStyle.Fill;
            //
            panel_po_tab1.Visible = false;
            panel_po_tab1.Dock = DockStyle.None;
            //
            panel_po_tab2_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_po_tab1_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void label119_Click(object sender, EventArgs e)
        {
            //
            panel_po_tab2.Visible = true;
            panel_po_tab2.Dock = DockStyle.Fill;
            //
            panel_po_tab1.Visible = false;
            panel_po_tab1.Dock = DockStyle.None;
            //
            panel_po_tab2_btn.BackgroundImage = Properties.Resources.tab_active;
            panel_po_tab1_btn.BackgroundImage = Properties.Resources.tab_inactive;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell.RowIndex == dataGridView1.RowCount - 1)
            {
                dataGridView1.AllowUserToDeleteRows = false;
            }
            else
            {
                dataGridView1.AllowUserToDeleteRows = true;
            }
        }

        private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            int indexOfTotalRow = dataGridView1.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView1.Rows[k].Cells[0].Value = k + 1;
                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[k].Cells[3].Value);
            }
            dataGridView1.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int indexOfTotalRow = dataGridView1.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[k].Cells[3].Value);
            }
            dataGridView1.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentCell.RowIndex == dataGridView6.RowCount - 1)
            {
                dataGridView6.AllowUserToDeleteRows = false;
            }
            else
            {
                dataGridView6.AllowUserToDeleteRows = true;
            }
        }

        private void dataGridView6_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            int indexOfTotalRow = dataGridView6.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                dataGridView6.Rows[k].Cells[0].Value = k + 1;
                totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[k].Cells[3].Value);
            }
            dataGridView6.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
        }

        private void dataGridView6_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int indexOfTotalRow = dataGridView6.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[k].Cells[3].Value);
            }
            dataGridView6.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
        }

        private void addRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(dataGridView1.RowCount==0)
            {
                dataGridView1.RowCount = 2;
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = dataGridView1.RowCount - 1;
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value = "0";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = "";
            }
            else
            {
                int rowcountOld = 0;
                rowcountOld = dataGridView1.RowCount;
                dataGridView1.RowCount = rowcountOld + 1;
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = dataGridView1.RowCount - 1;
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value = "0";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "";
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = "";
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                int k = dataGridView1.Rows.IndexOf(row);
                if (k % 2 == 0)
                {
                    row.DefaultCellStyle.BackColor = cl_even;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = cl_odd;
                }
                row.HeaderCell.Value = "";
            }

            int indexOfTotalRow = dataGridView1.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView1.Rows[k].Cells[3].Value);
            }
            dataGridView1.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
            dataGridView1.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
            dataGridView1.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.FirstDisplayedScrollingRowIndex + 1;
        }

        private void добавитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView6.RowCount == 0)
            {
                dataGridView6.RowCount = 2;
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[0].Value = dataGridView6.RowCount - 1;
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[3].Value = "0";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[1].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[2].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[4].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[5].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[6].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[7].Value = "";
            }
            else
            {
                int rowcountOld = 0;
                rowcountOld = dataGridView6.RowCount;
                dataGridView6.RowCount = rowcountOld + 1;
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[0].Value = dataGridView6.RowCount - 1;
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[3].Value = "0";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[1].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[2].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[4].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[5].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[6].Value = "";
                dataGridView6.Rows[dataGridView6.RowCount - 2].Cells[7].Value = "";
            }
            foreach (DataGridViewRow row in dataGridView6.Rows)
            {
                int k = dataGridView6.Rows.IndexOf(row);
                if (k % 2 == 0)
                {
                    row.DefaultCellStyle.BackColor = cl_even;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = cl_odd;
                }
                row.HeaderCell.Value = "";
            }

            int indexOfTotalRow = dataGridView6.RowCount - 1;
            int totalQty = 0;
            for (int k = 0; k < indexOfTotalRow; k++)
            {
                totalQty = totalQty + Convert.ToInt32(dataGridView6.Rows[k].Cells[3].Value);
            }
            dataGridView6.Rows[indexOfTotalRow].Cells[1].Value = "TOTAL: ";
            dataGridView6.Rows[indexOfTotalRow].Cells[3].Value = totalQty;
            dataGridView6.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
            dataGridView6.FirstDisplayedScrollingRowIndex = dataGridView6.FirstDisplayedScrollingRowIndex + 1;
        }

        private void dataGridView9_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridViewTextBoxEditingControl tb = (DataGridViewTextBoxEditingControl)e.Control;
            tb.KeyPress += new KeyPressEventHandler(dataGridView9_KeyPress);
            e.Control.KeyPress += new KeyPressEventHandler(dataGridView9_KeyPress);
        }

        private void dataGridView9_KeyPress(object sender, KeyPressEventArgs e)
        {
            int colIndex = dataGridView9.CurrentCell.ColumnIndex;
            if(colIndex>4 && colIndex<23)
            {
                string s = ".0123456789";
                if (s.IndexOf(e.KeyChar) >= 0 || e.KeyChar == (char)Keys.Back)
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            }
        }

        private void dataGridView12_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridViewTextBoxEditingControl tb = (DataGridViewTextBoxEditingControl)e.Control;
            tb.KeyPress += new KeyPressEventHandler(dataGridView12_KeyPress);
            e.Control.KeyPress += new KeyPressEventHandler(dataGridView12_KeyPress);
        }

        private void dataGridView12_KeyPress(object sender, KeyPressEventArgs e)
        {
            int colIndex = dataGridView12.CurrentCell.ColumnIndex;
            if (colIndex > 4 && colIndex < 23)
            {
                string s = ".0123456789";
                if (s.IndexOf(e.KeyChar) >= 0 || e.KeyChar == (char)Keys.Back)
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            }
        }

        private void dataGridView9_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //if (e.RowIndex == -1)
            //{
            //    Color c1 = cl_hc1;
            //    Color c2 = cl_hc2;
            //    Color c3 = cl_hc3;

            //    LinearGradientBrush br = new LinearGradientBrush(e.CellBounds, c1, c3, 90, true);
            //    ColorBlend cb = new ColorBlend();
            //    cb.Positions = new[] { 0, (float)0.5, 1 };
            //    cb.Colors = new[] { c1, c2, c3 };
            //    br.InterpolationColors = cb;

            //    e.Graphics.FillRectangle(br, e.CellBounds);
            //    e.PaintContent(e.ClipBounds);
            //    e.Handled = true;
            //}
        }

        private void SelectRootFilder(string folder)
        {
            try
            {
                //Создаём или перезаписываем существующий файл
                StreamWriter sw = File.CreateText(app_dir_temp + "root.txt");
                //Записываем текст в поток файла
                sw.WriteLine(folder);
                //Закрываем файл
                sw.Close();
                contextMenuStrip7.Items[1].Text = "Папка: " + folder;
            }
            catch { }
        }

        private void указатьРабочуюПапкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "Выберите рабочую папку в которой будут сохраняться все генерируемые документы";
            if (fd.ShowDialog() == DialogResult.OK) {
                string folder = fd.SelectedPath;
                SelectRootFilder(folder);
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void папкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string folder = contextMenuStrip7.Items[1].Text.Substring(7);
                System.Diagnostics.Process.Start(folder);
            }
            catch { }
        }

        // axad 22.07 start 
        private void dataGridView24_SelectionChanged(object sender, EventArgs e)
        {
            textBox37.Text = "";
            textBox38.Text = "";
            textBox39.Text = "";
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            int selected = 0;
            try
            {
                dataGridView14.RowCount = 0;
                selected = dataGridView13.CurrentCell.RowIndex;
                string id_contract = dataGridView13.Rows[selected].Cells[1].Value.ToString();
                selected = dataGridView24.CurrentCell.RowIndex;
                string version_contract = dataGridView24.Rows[selected].Cells[0].Value.ToString();
                GetContractItems(id_contract, version_contract);
            }
            catch { };
        }

        private void button54_Click(object sender, EventArgs e)
        {
            dataGridView14.AllowUserToDeleteRows = true;
            dataGridView14.ReadOnly = false;
            contextMenuStrip8.Enabled = true;
            contextMenuStrip8.Items[0].Enabled = true;
            dataGridView13.Enabled = false;
            dataGridView24.Enabled = false;
            button55.Enabled = true;
            button57.Enabled = true;
            button54.Enabled = false;
            textBox37.Text = "";
            textBox38.Text = "";
            textBox39.Text = "";
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (dataGridView14.RowCount == 0)
            {
                dataGridView14.RowCount = 2;
            }
            else
            {
                dataGridView14.RowCount = dataGridView14.RowCount + 1;
            }
            dataGridView14.Rows[dataGridView14.RowCount - 2].Cells[0].Value = dataGridView14.RowCount - 1;
            dataGridView14.Rows[dataGridView14.RowCount - 2].Cells[1].Value = "";
            dataGridView14.Rows[dataGridView14.RowCount - 2].Cells[2].Value = "";
            dataGridView14.Rows[dataGridView14.RowCount - 2].Cells[3].Value = 0;
            dataGridView14.Rows[dataGridView14.RowCount - 2].Cells[4].Value = 0;
            dataGridView14.Rows[dataGridView14.RowCount - 2].Cells[5].Value = 0;
            for (int i = 0; i < dataGridView14.RowCount - 1; i++)
            {
                if (i % 2 == 0)
                {
                    dataGridView14.Rows[i].Cells[0].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[1].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[2].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[3].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[4].Style.BackColor = cl_even;
                    dataGridView14.Rows[i].Cells[5].Style.BackColor = cl_even;
                }
                else
                {
                    dataGridView14.Rows[i].Cells[0].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[1].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[2].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[3].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[4].Style.BackColor = cl_odd;
                    dataGridView14.Rows[i].Cells[5].Style.BackColor = cl_odd;
                }
            }
            CountTotalForContract();
        }

        private void button55_Click(object sender, EventArgs e)
        {
            dataGridView14.AllowUserToDeleteRows = false;
            dataGridView14.ReadOnly = true;
            contextMenuStrip8.Enabled = false;
            contextMenuStrip8.Items[0].Enabled = false;
            dataGridView13.Enabled = true;
            dataGridView24.Enabled = true;
            button55.Enabled = false;
            button57.Enabled = false;
            button54.Enabled = true;
            dataGridView14.RowCount = 0;
        }

        private void dataGridView14_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CountTotalForContract();
        }
        public void CountTotalForContract()
        {
            int indexOfTotalRow = 0;
            indexOfTotalRow = dataGridView14.RowCount - 1;
            int qty = 0;
            double price = 0;
            double totalPrice = 0;
            for (int k = 0; k < dataGridView14.RowCount - 1; k++)
            {
                qty = qty + Convert.ToInt32(dataGridView14.Rows[k].Cells[3].Value.ToString());
                price = price + Math.Round(Convert.ToDouble(dataGridView14.Rows[k].Cells[4].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
                totalPrice = totalPrice + Math.Round(Convert.ToDouble(dataGridView14.Rows[k].Cells[5].Value.ToString(), new System.Globalization.CultureInfo("en-US")), 2, MidpointRounding.ToEven);
            }

            dataGridView14.Rows[indexOfTotalRow].Cells[2].Value = "TOTAL: ";
            dataGridView14.Rows[indexOfTotalRow].Cells[3].Value = qty;
            dataGridView14.Rows[indexOfTotalRow].Cells[4].Value = Math.Round(price, 2, MidpointRounding.ToEven).ToString("0.00");
            dataGridView14.Rows[indexOfTotalRow].Cells[5].Value = Math.Round(totalPrice, 2, MidpointRounding.ToEven).ToString("0.00");
            dataGridView14.Rows[indexOfTotalRow].DefaultCellStyle.BackColor = Color.LightSlateGray;
        }

        private void dataGridView14_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            CountTotalForContract();
            for(int i = 0; i<dataGridView14.RowCount - 1; i++)
            {
                dataGridView14.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void button57_Click(object sender, EventArgs e)
        {
            if(dataGridView14.RowCount>0)
            {
                if (checkBox11.Checked == false && checkBox12.Checked == false && checkBox13.Checked == false)
                {
                    MessageBox.Show("Please check the terms of payment", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (checkBox14.Checked == false && checkBox15.Checked == false)
                    {
                        MessageBox.Show("Please check the terms of delivery", "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        if (textBox37.Text.ToString().Length <= 0 || textBox38.Text.ToString().Length <= 0 || textBox39.Text.ToString().Length <= 0)
                        {
                            MessageBox.Show("Please fill the payment details", "Сообщение",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            SaveNewVersionOfContract();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Contract cannot be empty", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void SaveNewVersionOfContract()
        {
            if (comboBox18.Text != "")
            {
                string contract_date = "";
                string id_contract = "";
                string item_code = "";
                string name = "";
                string quantity = "";
                string unit_price = "";
                string amount_price = "";
                string delivery_point_air = "";
                string delivery_point_rw = "";

                string delivery_type = "";
                string cp_id = "";
                string cp_date = "";
                string c_order_id = "";
                string c_order_dt = "";
                string po_order_id = "";
                string po_order_dt = "";
                string cp_version = "";
                int v = 0;
                string customer = "";
                string amount = "";
                //
                string pay_before_percent = "";
                string pay_before_period = "";
                string delivery_period = "";
                string terms_of_payment = "";
                string terms_of_delivery = "";
                //
                pay_before_percent = textBox37.Text.ToString();
                pay_before_period = textBox38.Text.ToString();
                delivery_period = textBox39.Text.ToString();
                if (checkBox11.Checked == true)
                {
                    terms_of_payment = "1";
                }
                if (checkBox12.Checked == true)
                {
                    terms_of_payment = "2";
                }
                if (checkBox13.Checked == true)
                {
                    terms_of_payment = "3";
                }
                if (checkBox14.Checked == true)
                {
                    terms_of_delivery = "1";
                }
                if (checkBox15.Checked == true)
                {
                    terms_of_delivery = "2";
                }

                string year = Convert.ToString(dateTimePicker12.Value.Year);
                string month = Convert.ToString(dateTimePicker12.Value.Month);
                string day = Convert.ToString(dateTimePicker12.Value.Day);
                if (month.Length == 1) { month = "0" + month; }
                if (day.Length == 1) { day = "0" + day; }
                contract_date = day + "." + month + "." + year;
                v = dataGridView24.RowCount + 1;
                delivery_type = comboBox18.Text;

                int selected1 = dataGridView13.CurrentCell.RowIndex;
                id_contract = dataGridView13.Rows[selected1].Cells[1].Value.ToString();
                customer = dataGridView13.Rows[selected1].Cells[2].Value.ToString();
                amount = dataGridView13.Rows[selected1].Cells[3].Value.ToString();

                delivery_point_air = textBox41.Text;
                delivery_point_rw = textBox42.Text;

                string sql = "SELECT id_cp, cp_date, client_order_id, client_order_id_date, id_order, id_order_date, cp_version FROM contracts WHERE id_contract='" + id_contract + "' and version=1";
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                cp_id = reader.GetValue(0).ToString();
                                cp_date = reader.GetValue(1).ToString();
                                c_order_id = reader.GetValue(2).ToString();
                                c_order_dt = reader.GetValue(3).ToString();
                                po_order_id = reader.GetValue(4).ToString();
                                po_order_dt = reader.GetValue(5).ToString();
                                cp_version = reader.GetValue(6).ToString();
                            }
                        }
                    }
                    connection.Close();
                }

                SqlConnection connection2 = new SqlConnection(conString);
                connection2.Open();

                SqlCommand command2 = new SqlCommand();
                command2.Connection = connection2;
                command2.CommandType = CommandType.Text;
                command2.CommandText = "INSERT INTO contracts(id_contract, date, delivery_type, id_cp, cp_date, client_order_id, client_order_id_date, id_order, id_order_date, delivery_point_air, cp_version,  customer, amount, version, delivery_point_rw, pay_before_percent, pay_before_period, delivery_period, terms_of_payment, terms_of_delivery) VALUES(@idcont, @date, @delivery, @idcp, @cpdate, @clidorder, @clientorderdate, @idorder, @idorderdate, @dpoint_air, @cp_v, @c, @a, @version, @dpoint_rw, @paybefore, @paybefore_period, @delivery_period, @termspay, @termsdel)";
                command2.Parameters.AddWithValue("@idcont", id_contract);
                command2.Parameters.AddWithValue("@date", contract_date);
                command2.Parameters.AddWithValue("@delivery", delivery_type);
                command2.Parameters.AddWithValue("@idcp", cp_id);
                command2.Parameters.AddWithValue("@cpdate", cp_date);
                command2.Parameters.AddWithValue("@clidorder", c_order_id);
                command2.Parameters.AddWithValue("@clientorderdate", c_order_dt);
                command2.Parameters.AddWithValue("@idorder", po_order_id);
                command2.Parameters.AddWithValue("@idorderdate", po_order_dt);
                command2.Parameters.AddWithValue("@dpoint_air", delivery_point_air);
                command2.Parameters.AddWithValue("@cp_v", cp_version);
                command2.Parameters.AddWithValue("@c", customer);
                command2.Parameters.AddWithValue("@a", amount);
                command2.Parameters.AddWithValue("@version", v);
                command2.Parameters.AddWithValue("@dpoint_rw", delivery_point_rw);
                command2.Parameters.AddWithValue("@paybefore", pay_before_percent);
                command2.Parameters.AddWithValue("@paybefore_period", pay_before_period);
                command2.Parameters.AddWithValue("@delivery_period", delivery_period);
                command2.Parameters.AddWithValue("@termspay", terms_of_payment);
                command2.Parameters.AddWithValue("@termsdel", terms_of_delivery);

                command2.ExecuteNonQuery();

                int i = 0;
                for (i = 0; i < dataGridView14.Rows.Count - 1; i++)
                {
                    try { item_code = dataGridView14.Rows[i].Cells[1].Value.ToString(); } catch { item_code = ""; }
                    try { name = dataGridView14.Rows[i].Cells[2].Value.ToString(); } catch { name = ""; }
                    try { quantity = dataGridView14.Rows[i].Cells[3].Value.ToString(); } catch { quantity = ""; }
                    try { unit_price = dataGridView14.Rows[i].Cells[4].Value.ToString(); } catch { unit_price = ""; }
                    try { amount_price = dataGridView14.Rows[i].Cells[5].Value.ToString(); } catch { amount_price = ""; }

                    SqlCommand command3 = new SqlCommand();
                    command3.Connection = connection2;
                    command3.CommandType = CommandType.Text;
                    command3.CommandText = "INSERT INTO items_in_contract(id_contract, item_code, name, quantity, unit_price, amount_price, version) VALUES(@idcont, @itemcode, @name, @quantity, @unitprice, @amountprice, @v)";
                    command3.Parameters.AddWithValue("@idcont", id_contract);
                    command3.Parameters.AddWithValue("@itemcode", item_code);
                    command3.Parameters.AddWithValue("@name", name);
                    command3.Parameters.AddWithValue("@quantity", quantity);
                    command3.Parameters.AddWithValue("@unitprice", unit_price);
                    command3.Parameters.AddWithValue("@amountprice", amount_price);
                    command3.Parameters.AddWithValue("@v", v);
                    command3.ExecuteNonQuery();
                }

                connection2.Close();
                MessageBox.Show("контракт сохраненно.", "Done",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
                dataGridView14.AllowUserToDeleteRows = false;
                dataGridView14.ReadOnly = true;
                contextMenuStrip8.Enabled = false;
                contextMenuStrip8.Items[0].Enabled = false;
                dataGridView13.Enabled = true;
                dataGridView24.Enabled = true;
                button55.Enabled = false;
                button57.Enabled = false;
                button54.Enabled = true;

                int selected = 0;
                try
                {
                    dataGridView24.RowCount = 0;
                    dataGridView14.RowCount = 0;
                    selected = dataGridView13.CurrentCell.RowIndex;
                    string contract = dataGridView13.Rows[selected].Cells[1].Value.ToString();
                    GetContractVersions(contract);

                    dataGridView24.CurrentCell = dataGridView24.Rows[dataGridView24.RowCount - 1].Cells[0];
                    dataGridView24.CurrentCell.Selected = true;

                    textBox37.Text = "";
                    textBox38.Text = "";
                    textBox39.Text = "";
                    checkBox11.Checked = false;
                    checkBox12.Checked = false;
                    checkBox13.Checked = false;
                    checkBox14.Checked = false;
                    checkBox15.Checked = false;
                    int selected_1 = 0;
                    try
                    {
                        dataGridView14.RowCount = 0;
                        selected_1 = dataGridView13.CurrentCell.RowIndex;
                        string contract_id = dataGridView13.Rows[selected_1].Cells[1].Value.ToString();
                        selected_1 = dataGridView24.CurrentCell.RowIndex;
                        string version_contract = dataGridView24.Rows[selected_1].Cells[0].Value.ToString();
                        GetContractItems(contract_id, version_contract);
                    }
                    catch { };

                    if (dataGridView14.RowCount != 0)
                    {
                        int selected_generate = dataGridView13.CurrentCell.RowIndex;
                        string id_contract_generate = dataGridView13.Rows[selected_generate].Cells[1].Value.ToString();
                        string company_generate = dataGridView13.Rows[selected_generate].Cells[2].Value.ToString();

                        int selected_2_generate = dataGridView24.CurrentCell.RowIndex;
                        string delivery_point_generate = dataGridView24.Rows[selected_2_generate].Cells[2].Value.ToString();
                        string delivery_type_generate = dataGridView24.Rows[selected_2_generate].Cells[1].Value.ToString();
                        string data_generate = dataGridView24.Rows[selected_2_generate].Cells[3].Value.ToString();
                        string version_contract_generate = dataGridView24.Rows[selected_2_generate].Cells[0].Value.ToString();
                        generateContract(id_contract_generate, data_generate, delivery_type_generate, delivery_point_generate, company_generate, version_contract_generate);
                    }
                }
                catch { }
            }
            else
            {
                MessageBox.Show("Необходимо указать способ доставки.", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void checkBox11_Click(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                checkBox12.Checked = false;
            }
            if (checkBox13.Checked == true)
            {
                checkBox13.Checked = false;
            }
        }

        private void checkBox12_Click(object sender, EventArgs e)
        {
            if (checkBox11.Checked == true)
            {
                checkBox11.Checked = false;
            }
            if (checkBox13.Checked == true)
            {
                checkBox13.Checked = false;
            }
        }

        private void checkBox13_Click(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                checkBox12.Checked = false;
            }
            if (checkBox11.Checked == true)
            {
                checkBox11.Checked = false;
            }
        }

        private void checkBox14_Click(object sender, EventArgs e)
        {
            if (checkBox15.Checked == true)
            {
                checkBox15.Checked = false;
            }
        }

        private void checkBox15_Click(object sender, EventArgs e)
        {
            if (checkBox14.Checked == true)
            {
                checkBox14.Checked = false;
            }
        }

        private void checkBox17_Click(object sender, EventArgs e)
        {
            if (checkBox16.Checked == true)
            {
                checkBox16.Checked = false;
            }
        }

        private void checkBox16_Click(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true)
            {
                checkBox17.Checked = false;
            }
        }

        private void checkBox20_Click(object sender, EventArgs e)
        {
            if (checkBox19.Checked == true)
            {
                checkBox19.Checked = false;
            }
            if (checkBox18.Checked == true)
            {
                checkBox18.Checked = false;
            }
        }

        private void checkBox19_Click(object sender, EventArgs e)
        {
            if (checkBox20.Checked == true)
            {
                checkBox20.Checked = false;
            }
            if (checkBox18.Checked == true)
            {
                checkBox18.Checked = false;
            }
        }

        private void checkBox18_Click(object sender, EventArgs e)
        {
            if (checkBox19.Checked == true)
            {
                checkBox19.Checked = false;
            }
            if (checkBox20.Checked == true)
            {
                checkBox20.Checked = false;
            }
        }

    }   
        // axad 22.07 end
}
