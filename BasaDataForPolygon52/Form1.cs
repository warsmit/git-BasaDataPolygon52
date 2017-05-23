using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace BasaDataForPolygon52
{
    public partial class Form1 : Form
    {
        //для консоли
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
        //для консоли
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeConsole();
        
        int EndBD { set; get; }

        public List<Button> ListButSec = new List<Button>();
        public List<Button> ListButSubSec = new List<Button>();
        
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams result = base.CreateParams;
                result.ExStyle |= 0x02000000; // WS_EX_COMPOSITED
                return result;
            }
        }
        
        public Form1()
        {
            if (AllocConsole())
                Console.WriteLine("BD");

            InitializeComponent();
            
            SetDoubleBuffered(DataBase, true);

            //printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);

            DBExcel.Initialize();

            EndBD = DBExcel.EndDB();

            Console.Write("КОМИССИЯ = {0}\n", EndBD);

            labelData.Text = DateTime.Today.ToString("d");
            labelDataForOrder.Text = DateTime.Today.ToString("d");

            DG_Receipt.RowCount = 15;
            DG_Receipt.AllowUserToAddRows = false;

            DG_Receipt.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;


            //DG_Report.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dataGridView1_RowsAdded);
            //DG_Report.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.dataGridView1_RowsRemoved);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loadDB();

            loadButtonSection();
            locButton(ListButSec, panelButSec);

            loadButtonSubSection();

            comboBoxDiscount.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxDiscount.Items.Add("%");
            comboBoxDiscount.Items.Add("руб.");
            comboBoxDiscount.Text = "%";

            comboBoxSellers.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxSellers.Items.Add("Царев-Артур");
            comboBoxSellers.Items.Add("Артур-Марков");
            comboBoxSellers.Items.Add("Марков-Царев");

            timer1.Interval = 100;//Таймеры 
            timer1.Start();//Таймеры 

            labelSalesReceipt.Text = Properties.Settings.Default.NumberReceipt.ToString();

            resizeDGReport();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            labelTime.Text = DateTime.Now.ToString("HH:mm");
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            locButton(ListButSec, panelButSec);
            locButton(ListButSubSec, panelButSubSec);
            resizeDGReport();
        }



        public void SetDoubleBuffered(Control c, bool value)
        {
            PropertyInfo pi = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
            {
                pi.SetValue(c, value, null);

                MethodInfo mi = typeof(Control).GetMethod("SetStyle", BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.NonPublic);
                if (mi != null)
                {
                    mi.Invoke(c, new object[] { ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true });
                }

                mi = typeof(Control).GetMethod("UpdateStyles", BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.NonPublic);
                if (mi != null)
                {
                    mi.Invoke(c, null);
                }
            }
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            if (textBoxSearch.Text != string.Empty)
            {
                //сделать игнор русских англ букв
                Regex regex = new Regex("^" + textBoxSearch.Text + "$", RegexOptions.IgnoreCase);
                var temp = dataBasePolygon52.Where(s => regex.IsMatch(s.VendorCode));

                if (temp.Count() != 0)
                {
                    DataBase.DataSource = temp.ToArray();
                    DataBase.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    DataBase.Refresh();
                }
                else
                {
                    regex = new Regex(textBoxSearch.Text, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    temp = dataBasePolygon52.Where(s => regex.IsMatch(s.Name));
                    if (temp.Count() != 0)
                    {
                        DataBase.DataSource = temp.ToArray();
                        DataBase.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                        DataBase.Refresh();
                    }
                }
            }
        }

        private void DataBase_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                //Console.WriteLine(DataBase.Rows[e.RowIndex].Cells[1].Value.ToString());
                for(int i = 0; i < DG_Receipt.RowCount; i++)
                {
                    if(DG_Receipt.Rows[i].Cells[0].Value == null)
                    {
                        DG_Receipt.Rows[i].Cells[0].Value = DataBase.Rows[e.RowIndex].Cells[1].Value.ToString();
                        break;
                    }
                    else
                    {
                        if (DG_Receipt.Rows[i].Cells[0].Value.ToString() == DataBase.Rows[e.RowIndex].Cells[1].Value.ToString())
                            break;
                    }
                }
            }
        }

        private bool checkItemInReceipt(int row, int column)
        {
            if (DG_Receipt.Rows[row].Cells[column].Value != null)
                return true;
            else
                return false;
        }


        private bool getItemInReceipt(int row, int column)
        {
            if (DG_Receipt.Rows[row].Cells[column].Value != null)
                return true;
            else
                return false;
        }

        private void DG_Receipt_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0) return;

            int posVendorCode = 0;
            int posName = 1;
            int posCount = 2;
            int posPrice = 3;
            int posStock = 4;
            int posSum = 5;

            if (e.ColumnIndex == posVendorCode)
            {
                Regex regex = new Regex("^" + DG_Receipt[posVendorCode, e.RowIndex].Value.ToString() + "$", RegexOptions.IgnoreCase);

                var temp = dataBasePolygon52.Find(x => regex.IsMatch(x.VendorCode));

                if (temp != null)
                {
                    DG_Receipt[posName, e.RowIndex].Value = temp.Name;
                    DG_Receipt[posPrice, e.RowIndex].Value = temp.Price;
                    DG_Receipt[posStock, e.RowIndex].Value = temp.Stock;
                }
                else
                {
                    //DG_Receipt.Rows.RemoveAt(e.RowIndex);
                    DG_Receipt[posVendorCode, e.RowIndex].Value = string.Empty;
                    DG_Receipt[posName, e.RowIndex].Value = string.Empty;
                    DG_Receipt[posCount, e.RowIndex].Value = string.Empty;
                    DG_Receipt[posPrice, e.RowIndex].Value = string.Empty;
                    DG_Receipt[posStock, e.RowIndex].Value = string.Empty;
                    DG_Receipt[posSum, e.RowIndex].Value = string.Empty;
                }
            }
            else if (e.ColumnIndex == posCount)
            {
                if (DG_Receipt[posName, e.RowIndex].Value != null)
                {
                    //Console.WriteLine("posName " + posName + " e.RowIndex " + e.RowIndex);
                    //Console.WriteLine(DG_Receipt[posName, e.RowIndex].Value.ToString());
                    string count = DG_Receipt[posCount, e.RowIndex].Value.ToString();
                    if (Regex.IsMatch(count, @"^[0-9]{1,5}$"))
                    {
                        if (int.Parse(count) > 0)
                        {
                            int tempCount = int.Parse(DG_Receipt[posCount, e.RowIndex].Value.ToString());
                            int tempPrice = int.Parse(DG_Receipt[posPrice, e.RowIndex].Value.ToString());
                            int tempSum = tempCount * tempPrice;
                            DG_Receipt[posSum, e.RowIndex].Value = tempSum;
                            //totalSum += tempSum;
                        }
                        else
                        {
                            DG_Receipt[posCount, e.RowIndex].Value = string.Empty;
                            DG_Receipt[posSum, e.RowIndex].Value = string.Empty;
                        }
                    }
                    else
                        DG_Receipt[posCount, e.RowIndex].Value = string.Empty;

                    FinalSumWithDiscount();
                }
                else
                    DG_Receipt[posCount, e.RowIndex].Value = null;
            }   
        }

        int finalSum = 0;

        private void FinalSumWithDiscount()
        {
            labelDiscount.Text = string.Empty;
            labelTotal.Text = string.Empty;

            int posStock = 4;
            int posSum = 5;
            int sumForDiscount = 0;
            int sumForUndiscount = 0;

            for (int i = 0; i < DG_Receipt.RowCount; i++)
            {
                if (DG_Receipt.Rows[i].Cells[posSum].Value != null)
                {
                    if(DG_Receipt.Rows[i].Cells[posStock].Value.ToString() == string.Empty)
                        sumForDiscount += int.Parse(DG_Receipt.Rows[i].Cells[posSum].Value.ToString());
                    else
                        sumForUndiscount += int.Parse(DG_Receipt.Rows[i].Cells[posSum].Value.ToString());
                }
            }

            int discount = 0;
            if (comboBoxDiscount.SelectedIndex == 0)
            {
                if (textBoxDiscount.Text != string.Empty && sumForDiscount > 1000)
                {
                    discount = getPriceDiscount(sumForDiscount, int.Parse(textBoxDiscount.Text));
                    labelDiscount.Text = discount.ToString() + " р.";
                }
            }
            else if (comboBoxDiscount.SelectedIndex == 1)
            {
                if (int.Parse(textBoxDiscount.Text) >= sumForDiscount)
                {
                    MessageBox.Show(this, "Скидка больше суммы", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBoxDiscount.Text = "0";
                    discount = int.Parse(textBoxDiscount.Text);
                    labelDiscount.Text = discount.ToString() + " р.";
                }
                else
                {
                    discount = int.Parse(textBoxDiscount.Text);
                    labelDiscount.Text = discount.ToString() + " р.";
                }
            }

            finalSum = 0;
            finalSum = sumForDiscount - discount + sumForUndiscount;
            labelTotal.Text = finalSum.ToString() + " р.";
        }

        private void FinalSumUndiscounted()
        {
            labelDiscount.Text = string.Empty;
            labelTotal.Text = string.Empty;
            int sumForUndiscount = 0;
            int posSum = 5;

            for (int i = 0; i < DG_Receipt.RowCount; i++)
                if (DG_Receipt.Rows[i].Cells[posSum].Value != null)
                    sumForUndiscount += int.Parse(DG_Receipt.Rows[i].Cells[posSum].Value.ToString());

            finalSum = 0;
            finalSum = sumForUndiscount;
            labelTotal.Text = finalSum.ToString() + " р.";
        }

        private int getPriceDiscount(int price, int discount)
        {
            int MULTIPLE = 50;
            int newPrice = price * discount / 100;
            int remainder = newPrice % MULTIPLE;
            newPrice -= remainder;
            newPrice += (remainder >= MULTIPLE / 2) ? MULTIPLE : 0;
            return newPrice;
        }

        private void loadButtonSection()
        {
            //LINQ - получаем список разделов. По списку создаём кнопки для разделов
            foreach (var temp in dataBasePolygon52.Where(s => s.Section.Count() >= 1).Select(g => g.Section).Distinct())
            {
                CreateButtonSection(temp);
            }
        }

        private void loadButtonSubSection()
        {
            //LINQ - получаем список подразделов. По списку создаём кнопки для подразделов
            foreach (var temp in dataBasePolygon52.Where(s => s.SubSection.Count() >= 1).Select(g => g.SubSection).Distinct())
            {
                CreateButtonSubSection(temp);
            }
        }

        private void CreateButtonSection(string name)
        {
            Button button = new Button();

            sizeAndTextForButton(ref button, name);
            button.BackColor = Color.Silver;

            button.Click += (sender, args) =>
            {
                HideAllSubSection();
                IdentifierUseSection(button);
                var tempSection = dataBasePolygon52.Where(s => s.Section == button.Text);

                DataBase.DataSource = tempSection.ToArray();
                DataBase.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                DataBase.Refresh();

                //LINQ - получаем список подразделов. По списку отображаем соотвествующие кнопки
                foreach (var temp in tempSection.Where(s => s.SubSection.Count() >= 1).Select(g => g.SubSection).Distinct())
                {
                    ListButSubSec.Find(x => x.Text == temp).Show();
                }

                locButton(ListButSubSec, panelButSubSec); 
            };
            ListButSec.Add(button);
        }

        private void CreateButtonSubSection(string name)
        {
            Button button = new Button();

            sizeAndTextForButton(ref button, name);
            button.BackColor = Color.Silver;

            button.Click += (sender, args) =>
            {
                IdentifierUseSubSection(button);

                DataBase.DataSource = dataBasePolygon52.Where(s => s.SubSection == button.Text).ToArray();
            };

            button.Visible = false;

            ListButSubSec.Add(button);
        }

        private void IdentifierUseSection(Button butt)
        {
            try { ListButSec.Find(x => x.BackColor == Color.Red).BackColor = Color.Silver; } catch { }
            try { ListButSubSec.Find(x => x.BackColor == Color.Red).BackColor = Color.Silver; } catch { }
            butt.BackColor = Color.Red;
        }

        private void IdentifierUseSubSection(Button butt)
        {
            try { ListButSubSec.Find(x => x.BackColor == Color.Red).BackColor = Color.Silver; } catch { }
            butt.BackColor = Color.Red;
        }

        private void HideAllSubSection()
        {
            foreach (var subButton in ListButSubSec)
            {
                subButton.Visible = false;
            }
        }

        private void viewInDataBase(Item position)
        {
            DataBase.Rows.Add("В чек", position.Name, position.VendorCode, position.Count, position.Price, position.Stock, position.Brand);
        }

        private void locButton(List<Button> listButt, Panel panel)
        {
            Point locButton = new Point(0, 0);
            int lastButton = 0;

            for (int i = 0; i < listButt.Count(); i++)
            {
                if (!listButt[i].Visible)
                    continue;

                if (locButton.X + listButt[i].Width > panel.Width)
                {
                    locButton.X = 0;
                    locButton.Y += listButt[i].Height;
                }

                listButt[i].Location = locButton;
                locButton.X += listButt[i].Width;
                panel.Controls.Add(listButt[i]);
                lastButton = i;
            }

            panel.Height = listButt[lastButton].Location.Y + listButt[lastButton].Height;
        }

        private void sizeAndTextForButton(ref Button but, string name)
        {
            but.Font = new Font("Times New Roman", 14);

            using (Graphics cg = this.CreateGraphics())
            {
                SizeF sizeFond = cg.MeasureString(name, but.Font);
                but.Width = (int)sizeFond.Width + 20;
                but.Height = 40;
            }

            but.Text = name;
        }

        public List<Item> dataBasePolygon52 { set; get; }
        private void loadDB()
        {
            dataBasePolygon52 = new List<Item>();
            for (int i = 3; i < EndBD; i++)
            {
                dataBasePolygon52.Add(new Item(DBExcel.getVendorCode(i), DBExcel.getName(i), DBExcel.getCount(i), DBExcel.getPrice(i),DBExcel.getStock(i), DBExcel.getSection(i), DBExcel.getSubSection(i), DBExcel.getBrand(i), i));
            }
        }

        public class Item
        {
            public Item(string vendorCode, string name, int count, int price, string stock, string section, string subSection, string brand, int posInExcel)
            {
                VendorCode = vendorCode;
                Name = name;
                Count = count;
                Price = price;
                Stock = stock;
                Section = section;
                SubSection = subSection;
                Brand = brand;
                PosInExcel = posInExcel;
            }

            [DisplayName("Артикул")]
            public string VendorCode { set; get; }

            [DisplayName("Название")]
            public string Name { set; get; }

            [DisplayName("Количество")]
            public int Count { set; get; }

            [DisplayName("Цена")]
            public int Price { set; get; }

            [DisplayName("Акция")]
            public string Stock { set; get; }
            
            [DisplayName("Раздел")]
            public string Section { set; get; }

            [DisplayName("Подраздел")]
            public string SubSection { set; get; }

            [DisplayName("Брэнд")]
            public string Brand { set; get; }

            [DisplayName("Позиция")]
            public int PosInExcel { set; get; }
        }








        private void textBoxDiscount_TextChanged(object sender, EventArgs e)
        {
            if (textBoxDiscount.Text != string.Empty)
            {
                string pattern = string.Empty;
                if (comboBoxDiscount.SelectedIndex == 0)
                {
                    pattern = @"^[0-9]{1,2}$";
                }
                else if (comboBoxDiscount.SelectedIndex == 1)
                {
                    pattern = @"^[0-9]{1,6}$";
                }

                Regex regex = new Regex(pattern);
                if (regex.IsMatch(textBoxDiscount.Text))
                {
                    FinalSumWithDiscount();
                }
                else
                {
                    FinalSumUndiscounted();
                    MessageBox.Show(this, "Введите числовое значение", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBoxDiscount.Text = string.Empty;
                }
            }
            else
                FinalSumUndiscounted();
        }



        private void comboBoxDiscount_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxDiscount.Text = string.Empty;
            if (comboBoxDiscount.SelectedIndex == 0)
            {
                textBoxDiscount.MaxLength = 2;
            }
            else if (comboBoxDiscount.SelectedIndex == 1)
            {
                textBoxDiscount.MaxLength = 6;
            }
        }










        /*
        PrintDocument printDocument1 = new PrintDocument();
        PrintPreviewDialog previewdlg = new PrintPreviewDialog();
        Panel pannel = null;

        Bitmap MemoryImage;
        public void GetPrintArea(Panel pnl)
        {
            MemoryImage = new Bitmap(pnl.Width, pnl.Height);
            Rectangle rect = new Rectangle(0, 0, pnl.Width, pnl.Height);
            pnl.DrawToBitmap(MemoryImage, new Rectangle(0, 0, pnl.Width, pnl.Height));
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.DrawImage(MemoryImage, 0, 0);
            base.OnPaint(e);
        }
        void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Rectangle pagearea = e.PageBounds;
            e.Graphics.DrawImage(MemoryImage, (pagearea.Width / 2) - (pannel.Width / 2), pannel.Location.Y);
        }

        public void Print(Panel pnl)
        {
            pannel = pnl;
            GetPrintArea(pnl);
            previewdlg.Document = printDocument1;
            
            previewdlg.ShowDialog();
            
        }


            */





















        //очистка чека
        private void buttonClearReceipt_Click(object sender, EventArgs e)
        {
            clearReceipt();
        }

        private void clearReceipt()
        {
            DG_Receipt.Rows.Clear();
            DG_Receipt.RowCount = 14;
            textBoxDiscount.Text = string.Empty;
        }



        private void buttonPrintReceipt_Click(object sender, EventArgs e)
        {
            if (checkReceiptForNull())
            {
                AddReceiptInReport();
                clearReceipt();
                Properties.Settings.Default.NumberReceipt++;
                Properties.Settings.Default.Save();
                labelSalesReceipt.Text = Properties.Settings.Default.NumberReceipt.ToString();
            }
            else
            {
                MessageBox.Show(this, "Заполните столб количество", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //проверка, есть ли позиции в которых не указано количество
        private bool checkReceiptForNull()
        {
            for (int i = 0; i < DG_Receipt.RowCount; i++)
                if (DG_Receipt.Rows[i].Cells[1].Value != null && DG_Receipt.Rows[i].Cells[2].Value == null)
                    return false;
            return true;
        }

        private void AddReceiptInReport()
        {
            int countItemsReceipt = countItemsInReceipt();

            for (int i = 0; i < countItemsInReceipt(); i++)
            {
                DG_Report.Rows.Add();

                if(i == 0)
                    DG_Report.Rows[DG_Report.RowCount - 1].Cells[0].Value = labelSalesReceipt.Text;

                DG_Report.Rows[DG_Report.RowCount - 1].Cells[1].Value = DG_Receipt.Rows[i].Cells[0].Value;
                DG_Report.Rows[DG_Report.RowCount - 1].Cells[2].Value = DG_Receipt.Rows[i].Cells[1].Value;
                DG_Report.Rows[DG_Report.RowCount - 1].Cells[3].Value = DG_Receipt.Rows[i].Cells[2].Value;
                DG_Report.Rows[DG_Report.RowCount - 1].Cells[4].Value = DG_Receipt.Rows[i].Cells[3].Value;

                var tempPos = dataBasePolygon52.Find(x => x.VendorCode == DG_Receipt.Rows[i].Cells[0].Value.ToString());

                DBExcel.newCountForPosition(tempPos.PosInExcel, Convert.ToInt32(DG_Receipt.Rows[i].Cells[2].Value));

                tempPos.Count -= Convert.ToInt32(DG_Receipt.Rows[i].Cells[2].Value);
            }

            DBExcel.SaveExcelFile();
        }

        //количество итэмов в чеке
        private int countItemsInReceipt()
        {
            int i = 0;
            for (i = 0; i < DG_Receipt.RowCount; i++)
                if (DG_Receipt.Rows[i].Cells[0].Value == null)
                    return i;
            return i;
        }

        //ресайз отчёта
        private void resizeDGReport()
        {
            DG_Report.Height = this.Height - 150;
        }









        //удалить чек
        private void buttonDeleteReceipt_Click(object sender, EventArgs e)
        {
            deleteReceipt();
        }

        private void deleteReceipt()
        {
            string pattern = @"^[0-9]{1,10}$";
            Regex regex = new Regex(pattern);

            if (regex.IsMatch(textBoxNumReceipt.Text))
            {
                int startPosDel = 0;
                int countItemsDel = 0;

                if (getPosForDel(ref startPosDel, ref countItemsDel))
                {
                    for (int i = 0; i < countItemsDel; i++)
                        DG_Report.Rows.RemoveAt(startPosDel);
                }
                else
                    MessageBox.Show(this, "Не найден чек номером " + textBoxNumReceipt.Text, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show(this, "Введите числовое значение!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private bool getPosForDel(ref int start, ref int count)
        {
            for (int i = 0; i < DG_Report.RowCount; i++)
                if (check(DG_Report, i, 0))
                {
                    if (getInt(DG_Report, i, 0) == getInt(textBoxNumReceipt))
                    {
                        start = i;
                        count = 1;

                        if (DG_Report.RowCount > 1)
                        {
                            for (int j = start + 1; j < DG_Report.RowCount; j++)
                            {
                                if (check(DG_Report, j, 0))
                                {
                                    count = j;
                                    break;
                                }
                                else
                                    count = DG_Report.RowCount;
                            }
                        }
                        return true;
                    }
                }
            return false;
        }

        private bool check(DataGridView dg, int rows, int column)
        {
            return dg.Rows[rows].Cells[column].Value != null ? true : false;
        }

        private int getInt(DataGridView dg, int rows, int column)
        {
            return int.Parse(dg.Rows[rows].Cells[column].Value.ToString());
        }

        private int getInt(TextBox tb)
        {
            return int.Parse(tb.Text);
        }











        /*
         * автоудаление и добавление границы таблицы
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            // если в таблицу добавлена новая строка, то изменить высоту таблицы
            ChangeHeight();
        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            // если в таблице удалена строка, то изменить высоту таблицы
            ChangeHeight();
        }

        private void ChangeHeight()
        {
            // меняем высоту таблицу по высоте всех строк
            DG_Report.Height = DG_Report.Rows.GetRowsHeight(DataGridViewElementStates.Visible) +
                               DG_Report.ColumnHeadersHeight;
        }

        */














        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

                //NASSIM LOUCHANI
            }
        }

        //копировать всю датагриввив
        private void button2_Click(object sender, EventArgs e)
        {
            //копирование
            /*
            DG_Receipt.SelectAll();
            DataObject dataObj = DG_Receipt.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
            */
            //DG_Report.Rows.RemoveAt(0);


            DG_Report.Rows.Add();
            Console.WriteLine("DG_Report.RowCount " + DG_Report.RowCount);
            
            DG_Report.Rows[DG_Report.RowCount - 1].Cells[0].Value = "temp";
            DG_Report.Rows[DG_Report.RowCount - 1].Cells[1].Value = "temp";
            DG_Report.Rows[DG_Report.RowCount - 1].Cells[2].Value = "temp";
            DG_Report.Rows[DG_Report.RowCount - 1].Cells[3].Value = "temp";
            DG_Report.Rows[DG_Report.RowCount - 1].Cells[4].Value = "temp";
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            AddReceiptInReport();
            /*
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "temp.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(DataBase, sfd.FileName);
            }*/

        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadDB();
            DataBase.DataSource = dataBasePolygon52.Where(s => s.Section == "Оружие").ToArray();
            DataBase.Refresh();
        }
    }
}
