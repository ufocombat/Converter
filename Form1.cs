using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;//https://msdn.microsoft.com/ru-ru/library/hs600312(v=vs.110).aspx
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;//https://docs.microsoft.com/ru-ru/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
using Newtonsoft.Json;//https://stackoverflow.com/questions/4749639/deserializing-json-to-net-object-using-newtonsoft-or-linq-to-json-maybe
using System.Diagnostics;

namespace Converter
{
    public partial class Form1 : Form
    {
        Setup setup;
        List<Item> items = new List<Item>();
        List<Column> columns = new List<Column>();
        StringBuilder r = new StringBuilder();
        String ExcelFimeName = String.Empty;

        static String DirForms = Application.StartupPath + @"\Формы";
        static String DirReports = Application.StartupPath + @"\Отчеты";
        static String SetupFileName = Application.StartupPath + @"\setup.json";
        static String ConfigFileName = Application.StartupPath + @"\config.json";

        DataTable dt;

        private static String Normal(int value)
        {
            String work = value.ToString();
            while (work.Length < 3) work = "0" + work;
            return work;
        }

        private void Start()
        {
            toolFileName.Text = "Откройте файл Книги РСП (Ctl+O)";
            items.Clear();
            r.Clear();
            rt.Text = String.Empty;
        }

        public Form1()
        {
            InitializeComponent();

            Directory.CreateDirectory(DirForms);
            Directory.CreateDirectory(DirReports);

            comDistrict.Items.Clear();
            comDistrict.Items.Add("(все)");
            comDistrict.Items.Add("м");
            comDistrict.Items.Add("с");

            Start();

            try
            {
                setup = JsonConvert.DeserializeObject<Setup>(File.ReadAllText(SetupFileName));
            }
            catch
            {
                rt.Text = "Поврежден файл настройки.";
            }

            if (setup == null)
            {
                setup = new Setup();
                File.WriteAllText(SetupFileName, JsonConvert.SerializeObject(setup, Formatting.Indented));
            }

            comDistrict.Text = setup.District;

            if (File.Exists(ConfigFileName))
                try
                {
                    columns = JsonConvert.DeserializeObject<List<Column>>(File.ReadAllText(ConfigFileName));
                }
                catch
                {
                    rt.Text = "Поврежден файл конфигурации. ";
                }

            if (columns.Count == 0)
            {
                columns.Add(new Column("Статья УК", 9) { Key = true });
                columns.Add(new Column("Присоеденино", 10));
                File.WriteAllText(ConfigFileName, JsonConvert.SerializeObject(columns, Formatting.Indented));
            }
        }


        private void ReadBook()
        {
            StreamWriter logFile = null;

            if (setup.Trace != String.Empty)
            {
                logFile = new StreamWriter(Application.StartupPath + "\\log.txt", false, Encoding.UTF8);
                logFile.WriteLine(String.Format("Трассировка статьи {0}", setup.Trace));
            }

            int lines = 0;
            int i = 0;

            String distr = comDistrict.Text.ToLower();
            Boolean all = "(все)" == distr;

            for (i = setup.FromLineKRSP - 1; i < dt.Rows.Count; i++)
            {
                if (all || (dt.Rows[i][2].ToString() == distr))
                {
                    var item = new Item() { id = dt.Rows[i][5].ToString() };

                    item.dataRow = dt.Rows[i];



                    for (int n = 1; n < 13; n++) if (item.Is(n)) item.c[n]++;

                    var f = false;
                    lines++;


                    foreach (Item ii in items)
                    {
                        f = ii.id == item.id;
                        if (f)
                        {
                            for (int n = 0; n < 13; n++) ii.c[n] = ii.c[n] + item.c[n];
                            break;
                        }
                    }

                    if (!f) items.Add(item);

                    if ((logFile != null) && (item.id == setup.Trace))
                    {
                        logFile.Write(String.Format("СтрКниги[{0}],<{1}>;", Normal(i+1), dt.Rows[i][5].ToString()));
                        if (item.c[1]>0) logFile.Write("Присоеденено;");
                        logFile.WriteLine();
                    }
                }
            }

            r.AppendLine(String.Format("Обработано {0} строк", lines));
            if (logFile != null) logFile.Close();
        }

        private void OpenKRSP_Click(object sender, EventArgs e)
        {
            modeFile.Enabled = false;
            toolFileName.Text = String.Empty;

            //try
            {
                using (OpenFileDialog of = new OpenFileDialog() { Filter = "Файлы Excel|*.xls;*.xlsx", ValidateNames = true })
                {

                    if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        r.Clear();
                        r.AppendLine("Открыт файл: " + of.FileName);
                        r.AppendLine("Район: " + comDistrict.Text);

                        items.Clear();

                        ExcelFimeName = of.FileName;
                        FileStream stream;

                        try
                        {
                            stream = File.Open(ExcelFimeName, FileMode.Open, FileAccess.Read);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Не удалось открыть указанный файл. Если файл открыт в другой программе закроете его и повторите операцию.");
                            modeFile.Enabled = false;
                            return;
                        }

                        IExcelDataReader excelReader;

                        var file = new FileInfo(ExcelFimeName);
                        if (file.Extension.Equals(".xls"))
                            excelReader = ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(stream);
                        else if (file.Extension.Equals(".xlsx"))
                            excelReader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream);
                        else
                            throw new Exception("Данный формат файла не поддерживается.");

                        DataSet result = excelReader.AsDataSet();
                        //excelReader.IsFirstRowAsColumnNames = true;

                        if (dt != null) dt.Clear();
                        dt = result.Tables[0];

                        excelReader.Close();
                        stream.Close();

                        ReadBook();

                        using (StreamWriter writer = new StreamWriter(DirReports + "\\summary.csv", false, Encoding.UTF8))
                        {
                            writer.WriteLine("Статья УК;Количество;Присоедено (16=СЕ);Возбуждено (16=УД);Отказано (16=отказано);По пп 1, 2 ч. ст 24 (18=1|2);По пп 3 ч. ст 24 (18=3);По пп 4 ч. ст 24 (18=4);Передано в суд (16=под-ти|пор-ти);16=по под-ти;16=по пор-ти;Выезд (15=15|да|в др. суб.);Выезд 3-10 сут (19= <=10);Выезд больше 10 сут (19= >10)");

                            foreach (Item itemm in items)
                            {
                                writer.Write(String.Format("{0};", itemm.id));
                                for (int j = 0; j < 13; j++) writer.Write(String.Format("{0};", itemm.c[j]));
                                writer.WriteLine();
                            }
                        }

                        rt.Text = r.ToString();

                        OpenForm713();
                    }
                }
            }
            /*            catch e
                        {
                            ShowDialog()
                        }*/
            modeFile.Enabled = true;
        }
        static void DisplayInExcel()
        {
            var excelApp = new Excel.Application();

            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
        }
        private void OpenForm713()
        {

            r.AppendLine(String.Empty);
            r.AppendLine("Заполнение формы 713");
            r.AppendLine("Не закрывайте программу до окончания выгрузки (100%)");
            rt.Text = r.ToString();
            Application.DoEvents();

            String ExcelFileName;
            int i = 1;

            do
            {
                ExcelFileName = DirReports + String.Format("\\{0}_713.xlsx", Normal(i++));
            } while (File.Exists(ExcelFileName));

            File.Copy(DirForms + "\\713.xlsx", ExcelFileName);
            //DisplayInExcel();

            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(ExcelFileName);
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;


            for (int row = setup.FromLine713; row <= setup.ToLine713; row++)
            {
                toolFileName.Text = String.Format("{0}%", (int)row * 100 / setup.ToLine713);
                Application.DoEvents();

                try
                {
                    if (workSheet.Cells[row, "H"].Value2 != null)
                    {
                        String[] elements = Regex.Split(workSheet.Cells[row, "H"].Value.ToString(), @";\s*");
                        foreach (var element in elements)
                            foreach (Item item in items)
                            {

                                if (element.IndexOf(item.id) >= 0)
                                {
                                    try
                                    {
                                        var oRange = (Excel.Range)workSheet.Cells[row, 8];
                                        oRange.Interior.Color = System.Drawing.Color.LightGreen;
                                    }
                                    catch
                                    {
                                    }

                                    for (int j = 0; j < 13; j++)
                                    {
                                        int jj = j + 9;
                                        if (workSheet.Cells[row, jj].Value2 == null)
                                            workSheet.Cells[row, jj].Value = item.c[j];
                                        else
                                            workSheet.Cells[row, jj].Value = workSheet.Cells[row, jj].Value + item.c[j];
                                    }

                                    items.Remove(item);
                                    break;
                                }
                            }
                    }
                }
                catch
                {
                    var oRange = (Excel.Range)workSheet.Cells[row, "H"];
                    oRange.Interior.Color = System.Drawing.Color.Red;

                    r.Append(String.Format("Ошибка в строке формы {0}", row));
                }
            }

            r.AppendLine("Заполнение завершено.");
            if (items.Count > 0)
            {
                r.AppendLine(String.Empty);
                r.AppendLine("В н и м а н и е !");
                r.AppendLine("Не внесены статьи:");
                foreach (Item item in items)
                {
                    r.AppendLine(String.Format("Статья {0}", item.id));
                }
            }
            toolFileName.Text = String.Empty;
            rt.Text = r.ToString();
        }
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.ShowDialog(this);
        }
        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Start();
        }

        private void форма713ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(DirForms + @"\713.xlsx");
        }

        private void инструкцияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f = new Form3();
            f.ShowDialog(this);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            File.WriteAllText(SetupFileName, JsonConvert.SerializeObject(setup, Formatting.Indented));
        }

        private void comDistrict_TextChanged(object sender, EventArgs e)
        {
            setup.District = comDistrict.Text;
        }
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void видеоИнструкцииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.youtube.com/channel/UCfsezwUlKm_BAoHoz9uGH1A?view_as=subscriber");
        }
    }
}
