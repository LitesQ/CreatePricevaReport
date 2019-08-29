using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Exel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"parser\parser.bat");
            }
            catch (System.ComponentModel.Win32Exception) { Text = "ОШИБКА"; Status.Text = "В папке с программой должна быть папка parser\nПерезапустите программу" ; Action3.BackColor = Color.Gray; Action3.Enabled = false; Action3.Text = "Ошибка"; }
        }

        private void Action3_Click(object sender, EventArgs e)
        {
            Action3.BackColor = Color.Gray;
            Action3.Text = "Выполнение";
            Status.Text = "";
            Action3.Enabled = false;
            Text = "Отчет по ценам";

            openFileDialog1.InitialDirectory = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");
            openFileDialog1.ShowDialog();

            string PathToSourceResult = openFileDialog1.FileName;
            if (PathToSourceResult.Length == 0)
            {
                Text = "ОШИБКА";
                Action3.Text = "БКК";
                Status.Text = "Вы не выбрали файл";
                Action3.Enabled = true;
                Action3.BackColor = Color.DarkRed;
                return;
            }
            FileInfo finfo = new FileInfo(PathToSourceResult);

            if (!(finfo.Extension == ".xlsx" || finfo.Extension == ".xls"))
            {
                Text = "ОШИБКА";
                Action3.Text = "БКК";
                Status.Text = "Выбранный файл должен быть формата .xlsx или .xls";
                Action3.Enabled = true;
                Action3.BackColor = Color.DarkRed;
                return;
            }

            DateTime CurrentDate = new DateTime();
            CurrentDate = DateTime.Now;

            string LinksName;
            string ResultFileName;
            string[] Config;
            string[] ConfigSplitter = null;

            try
            {
                Config = File.ReadAllLines(Environment.CurrentDirectory + @"\config.ini");
            }
            catch (FileNotFoundException) { Text = "ОШИБКА"; Status.Text = "Файл config.ini не найден. Он должен быть в одной папке с программой."; Action3.BackColor = Color.DarkRed; Action3.Text = "БКК"; Action3.Enabled = true; return; }

            ConfigSplitter = Config[0].Split('=');
            LinksName = ConfigSplitter[1];
            ConfigSplitter = Config[1].Split('=');
            ResultFileName = ConfigSplitter[1];

            if (File.Exists(Environment.CurrentDirectory + @"\out\" + ResultFileName + "_" + CurrentDate.ToShortDateString() + finfo.Extension))
            {
                Text = "ОШИБКА";
                Action3.Text = "Ошибка";
                Status.Text = "Файл с названием " + ResultFileName + "_" + CurrentDate.ToShortDateString() + finfo.Extension + " в папке Out уже существует.";
                Action3.BackColor = Color.DarkRed;
                Action3.Enabled = true;
                return;
            }

            Process[] pname = Process.GetProcessesByName("php");
            Process[] pname1 = Process.GetProcessesByName("cmd");
            if (pname.Length != 0 && pname1.Length != 0)
            {
                Process proc = pname[0];
                Process proc1 = pname1[0];

                Text = "Отчет по ценам - Создание Base.csv";

                proc.WaitForExit();
                proc1.WaitForExit();
            }
            Text = "Отчет по ценам - Base.csv создан";

            string[] Dots = { ".", "..", "..." };
            int DotsCounter = 0;

            Excel.Application Links = new Excel.Application();
            try
            {
                Links.Workbooks.Open(Environment.CurrentDirectory + @"\in\" + LinksName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Text = "ОШИБКА";
                Action3.Text = "БКК";
                Status.Text = "Не найден файл " + LinksName + " с готовыми связками в папке in";
                Action3.Enabled = true;
                Action3.BackColor = Color.DarkRed;
                return;
            }
            Excel.Worksheet SheetLinks = (Excel.Worksheet)Links.Worksheets.get_Item(1);

            Excel.Range RangeGoodsID;
            Excel.Range xlRange = SheetLinks.UsedRange;
            string[] GoodsID = new string[xlRange.Rows.Count];

            for (int i = 0; i < xlRange.Rows.Count - 1; i++)
            {
                Application.DoEvents();
                RangeGoodsID = SheetLinks.Cells[i + 2, 3] as Excel.Range;
                GoodsID[i] = RangeGoodsID.Value2.ToString();
                Text = "Отчет по ценам - Считывание файла " + LinksName + Dots[DotsCounter];
                if (DotsCounter < 2)
                    DotsCounter++;
                else
                    DotsCounter = 0;
                Application.DoEvents();
            }

            Excel.Range RangeVendorCode;
            string[] VendorCodeArray = new string[xlRange.Rows.Count];

            for (int i = 0; i < xlRange.Rows.Count - 1; i++)
            {
                Application.DoEvents();
                RangeVendorCode = SheetLinks.Cells[i + 2, 11] as Excel.Range;
                VendorCodeArray[i] = RangeVendorCode.Value2.ToString();
                Text = "Отчет по ценам - Считывание файла " + LinksName + Dots[DotsCounter];
                if (DotsCounter < 2)
                    DotsCounter++;
                else
                    DotsCounter = 0;
                Application.DoEvents();
            }

            Excel.Application Result = new Excel.Application();
            Result.Workbooks.Open(PathToSourceResult,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
            Excel.Worksheet SheetResult = (Excel.Worksheet)Result.Worksheets.get_Item(1);

            string[] BaseValues = File.ReadAllLines(Environment.CurrentDirectory + @"\in\base.csv", System.Text.Encoding.GetEncoding(1251));

            string[] splitter = null;
            bool SellerExist = false;
            Excel.Range RangeResult;
            Excel.Range UsedRangeResilt;

            string IDinBase;
            string VendorCode;
            bool VendorCodeMatched = false;
            int CounterVC = 2;
            int ColumnNumber = 0;

            Dictionary<string, int> NewVendorCodes = new Dictionary<string, int>();
            bool CheckinNewVendorCodes = false;

            for (int i = 1; i < BaseValues.Length; i++)
            {
                Application.DoEvents();
                splitter = BaseValues[i].Split(';');
                Status.Text = "Текущая обрабатываемая строка: " + i + "\n" + "Всего строк на обработку: " + BaseValues.Length;
                for (int k = 3; k < splitter.Length; k += 4)// читаем селлеров
                {
                    Application.DoEvents();
                    UsedRangeResilt = SheetResult.UsedRange;
                    Text = "Отчет по ценам - Генерация отчета" + Dots[DotsCounter];     //Декор
                    if (DotsCounter < 2)
                        DotsCounter++;
                    else
                        DotsCounter = 0;                                                                    //
                    Application.DoEvents();
                    if (splitter[k] != "Bezant" && splitter[k] != "") // проверка селлера
                    {
                        for (int j = 13; j < UsedRangeResilt.Columns.Count + 1; j++) // проверка совпадения селлера
                        {
                            Application.DoEvents();
                            RangeResult = SheetResult.Cells[1, j] as Excel.Range;

                            if (splitter[k] == RangeResult.Value2.ToString())
                            {
                                SellerExist = true;
                                ColumnNumber = j;
                                break;
                            }
                            Application.DoEvents();
                        }// j
                        Application.DoEvents();
                        if (SellerExist == false)
                        {
                            SheetResult.Cells[1, UsedRangeResilt.Columns.Count + 1] = splitter[k];
                            RangeResult = SheetResult.Cells[1, UsedRangeResilt.Columns.Count + 1] as Excel.Range;
                            RangeResult.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xEA, 0xF2, 0xDD));
                            RangeResult.Cells.Font.Size = 10;
                            IDinBase = splitter[0];

                            for (int h = 0; h < GoodsID.Length; h++)
                            {
                                Application.DoEvents();
                                if (IDinBase == GoodsID[h])
                                {
                                    VendorCode = VendorCodeArray[h];

                                    while (CounterVC < UsedRangeResilt.Rows.Count + 1)// поиск артикула в результе
                                    {
                                        Application.DoEvents();
                                        RangeResult = SheetResult.Cells[CounterVC, 1] as Excel.Range;
                                        try
                                        {
                                            if (VendorCode == RangeResult.Value2.ToString())
                                            {
                                                SheetResult.Cells[CounterVC, UsedRangeResilt.Columns.Count + 1] = splitter[k + 1];
                                                VendorCodeMatched = true;
                                                CheckinNewVendorCodes = true;
                                                break;
                                            }
                                        }
                                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                                        {
                                            if (VendorCode == String.Empty)
                                            {
                                                SheetResult.Cells[CounterVC, UsedRangeResilt.Columns.Count + 1] = splitter[k + 1];
                                                VendorCodeMatched = true;
                                                CheckinNewVendorCodes = true;
                                                break;
                                            }
                                        }
                                        CounterVC++;
                                        Application.DoEvents();
                                    }
                                    if (CheckinNewVendorCodes == false)
                                    {
                                        foreach (KeyValuePair<string, int> key in NewVendorCodes)
                                        {
                                            Application.DoEvents();
                                            if (VendorCode == key.Key)
                                            {
                                                SheetResult.Cells[key.Value, UsedRangeResilt.Columns.Count + 1] = splitter[k + 1];
                                                VendorCodeMatched = true;
                                                break;
                                            }
                                            Application.DoEvents();
                                        }
                                    }
                                    Application.DoEvents();
                                    if (VendorCodeMatched == false)// Не нашел артукул в результе
                                    {
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, 1] = VendorCode;
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, 5] = splitter[1];
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, 7] = splitter[2];
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, UsedRangeResilt.Columns.Count + 1] = splitter[k + 1];
                                        NewVendorCodes.Add(VendorCode, UsedRangeResilt.Rows.Count + 1);
                                    }
                                    Application.DoEvents();
                                    CheckinNewVendorCodes = false;
                                    VendorCodeMatched = false;
                                    CounterVC = 2;
                                    break;
                                }// IDinBase == GoodsID[h]
                                Application.DoEvents();
                            } // h < GoodsID.Length
                            Application.DoEvents();
                        }
                        else if (SellerExist == true)
                        {
                            IDinBase = splitter[0];

                            for (int h = 0; h < GoodsID.Length; h++)
                            {
                                Application.DoEvents();
                                if (IDinBase == GoodsID[h])
                                {
                                    VendorCode = VendorCodeArray[h];

                                    while (CounterVC < UsedRangeResilt.Rows.Count + 1)
                                    {
                                        Application.DoEvents();
                                        RangeResult = SheetResult.Cells[CounterVC, 1] as Excel.Range;
                                        try
                                        {
                                            if (VendorCode == RangeResult.Value2.ToString())
                                            {
                                                SheetResult.Cells[CounterVC, ColumnNumber] = splitter[k + 1];
                                                VendorCodeMatched = true;
                                                CheckinNewVendorCodes = true;
                                                break;
                                            }
                                        }
                                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                                        {
                                            if (VendorCode == String.Empty)
                                            {
                                                SheetResult.Cells[CounterVC, ColumnNumber] = splitter[k + 1];
                                                VendorCodeMatched = true;
                                                CheckinNewVendorCodes = true;
                                                break;
                                            }
                                        }
                                        CounterVC++;
                                        Application.DoEvents();
                                    }
                                    if (CheckinNewVendorCodes == false)
                                    {
                                        foreach (KeyValuePair<string, int> key in NewVendorCodes)
                                        {
                                            Application.DoEvents();
                                            if (VendorCode == key.Key)
                                            {
                                                SheetResult.Cells[key.Value, ColumnNumber] = splitter[k + 1];
                                                VendorCodeMatched = true;
                                                break;
                                            }
                                            Application.DoEvents();
                                        }
                                    }
                                    if (VendorCodeMatched == false)// не нашел артикул
                                    {
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, 1] = VendorCode;
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, 5] = splitter[1];
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, 7] = splitter[2];
                                        SheetResult.Cells[UsedRangeResilt.Rows.Count + 1, ColumnNumber] = splitter[k + 1];
                                        NewVendorCodes.Add(VendorCode, UsedRangeResilt.Rows.Count + 1);
                                    }
                                    Application.DoEvents();
                                    CheckinNewVendorCodes = false;
                                    VendorCodeMatched = false;
                                    CounterVC = 2;
                                    break;
                                }// IDinBase == GoodsID[h]
                                Application.DoEvents();
                            }// h < GoodsID.Length
                            Application.DoEvents();
                        }
                        SellerExist = false;
                        Application.DoEvents();
                    }// if seller != "Bezant" && seller != ""
                    Application.DoEvents();
                }// k
                Application.DoEvents();
            }// i

            Links.Quit();
            Result.Application.ActiveWorkbook.SaveAs(Environment.CurrentDirectory + @"\out\" + ResultFileName + "_" + CurrentDate.ToShortDateString() + finfo.Extension, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Result.Quit();
            Process.Start("explorer", Environment.CurrentDirectory + @"\out\");
            Action3.BackColor = Color.DarkRed;
            Action3.Enabled = true;
            Action3.Text = "БКК";
            Text = "Отчет по ценам - Готово!";
            Status.Text = "Итоговый документ находится в папке out с названием " + ResultFileName + "_" + CurrentDate.ToShortDateString() + finfo.Extension;

        }
    }
}

