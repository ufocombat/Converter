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

            Start();

            if (File.Exists(SetupFileName))
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

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ReadBook()
        {
            int lines = 0;
            int i = 0;

            for (i = 4; i < dt.Rows.Count; i++)
            {
                var item = new Item()
                {
                    id = dt.Rows[i][5].ToString(),
                   // Solution = dt.Rows[i][15].ToString(),
                    Issue = dt.Rows[i][17].ToString(),
                    Outgo = dt.Rows[i][14].ToString(),
                    OutgoT = dt.Rows[i][18].ToString()
                };

                item.dataRow = dt.Rows[i];

                for (int n=1;n<13;n++) if (item.Is(n)) item.c[n]++;

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


            }

            r.AppendLine(String.Format("Обработано {0}  строк", lines));
            r.AppendLine(String.Format("Последняя строка {0}", i));
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

                        items.Clear();

                        ExcelFimeName = of.FileName;

                        FileStream stream = File.Open(ExcelFimeName, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelReader;

                        var file = new FileInfo(ExcelFimeName);
                        if (file.Extension.Equals(".xls"))
                            excelReader = ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(stream);
                        else if (file.Extension.Equals(".xlsx"))
                            excelReader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream);
                        else
                            throw new Exception("Данный формат файла не поддерживается. Обратитесь к разработчику.");

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

                            foreach (Item ii in items)
                            {
                                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13}",
                                    ii.id,
                                    ii.c[0],
                                    ii.c[1],
                                    ii.c[2],
                                    ii.c[3],
                                    ii.c[4],
                                    ii.c[5],
                                    ii.c[6],
                                    ii.c[7],
                                    ii.c[8],
                                    ii.c[9],
                                    ii.c[10],
                                    ii.c[11],
                                    ii.c[12]
                                   ));
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

                                    for (int j = 0; j < 12; j++)
                                    {
                                        int jj = j + 9;
                                        if (workSheet.Cells[row, jj].Value2 == null)
                                            workSheet.Cells[row, jj].Value = item.c[j];
                                        else
                                            workSheet.Cells[row, jj].Value = workSheet.Cells[row, jj].Value + item.c[j];

                                        /*workSheet.Cells[row, 10].Value = item.c[2];
                                        workSheet.Cells[row, 11].Value = item.c[3];
                                        workSheet.Cells[row, 12].Value = item.c[4];
                                        workSheet.Cells[row, 13].Value = item.c[5];
                                        workSheet.Cells[row, 14].Value = item.c[6];
                                        workSheet.Cells[row, 15].Value = item.c[7];
                                        workSheet.Cells[row, 16].Value = item.c[8];
                                        workSheet.Cells[row, 17].Value = item.c[9];
                                        workSheet.Cells[row, 18].Value = item.c[10];
                                        workSheet.Cells[row, 19].Value = item.c[11];
                                        workSheet.Cells[row, 20].Value = item.c[12];*/
                                    }

                                    //  r.AppendLine(String.Format("Статья {0} внесена в отчет", item.id));

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
    }
}
