using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace Converter
{
    class Item
    {
        public int[] c = { 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        public DataRow dataRow;

//        public String Solution = String.Empty;
        public String Issue = String.Empty;
        public String Outgo = String.Empty;
        public String OutgoT = String.Empty;

        public string id { get; set; }
        public String trace = String.Empty;

        public Boolean Is(int n)
        {
            switch (n)
            {
                case 1: return dataRow[15].ToString().ToUpper() == "СЕ";
                case 2: return dataRow[15].ToString().ToUpper() == "УД";
                case 3: return dataRow[15].ToString().ToLower() == "отказано";
                case 4:
                    {
                        string[] array = { "1", "2" };
                        return array.Contains(Issue);
                    }
                case 5: return Issue == "3";
                case 6: return Issue == "4";
                case 7:
                    {
                        string[] array = { "по под-ти", "по тер-ти" };
                        return array.Contains(dataRow[15].ToString().ToLower());
                    }
                case 8:
                    {
                        string[] array = { "по под-ти" };
                        return array.Contains(dataRow[15].ToString().ToLower());
                    }
                case 9:
                    {
                        string[] array = { "по тер-ти" };
                        return array.Contains(dataRow[15].ToString().ToLower());
                    }
                case 10:
                    {
                        string[] array = { "в др. суб.", "да", "15" };
                        return array.Contains(Outgo);
                    }
                case 11:
                    {
                        string[] array = { "3", "10" };
                        return array.Contains(OutgoT);
                    }
                case 12:
                    {
                        string[] array = { "19", "30" };
                        return array.Contains(OutgoT);
                    }
                default:
                    return false;
            }
        }
    }
}

