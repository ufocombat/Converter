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
        public string id { get; set; }
        

        String Cell(int n)
        {
            char[] charsToTrim = { ' ' };
            return dataRow[n-1].ToString().Trim(charsToTrim);
        }
        int CellInt(int n)
        {
            if (dataRow[n - 1] != DBNull.Value)
            {
                try
                {
                    return Convert.ToInt32(dataRow[n - 1]);
                }
                catch
                {
                    return -1;
                }
            }
            else
                return -1;
        }


public Boolean Is(int n)
        {
            

            switch (n)
            {
                case 1: return Cell(16).ToUpper() == "СЕ";
                case 2: return Cell(16).ToUpper() == "УД";
                case 3: return Cell(16).ToLower() == "отказано";
                case 4:
                    {
                        string[] array = { "1", "2" };
                        return array.Contains(Cell(18));
                    }
                case 5: return Cell(18) == "3";
                case 6: return Cell(18) == "4";
                case 7:
                    {
                        string[] array = { "по под-ти", "по тер-ти" };
                        return array.Contains(Cell(16).ToLower());
                    }
                case 8:
                    {
                        string[] array = { "по под-ти" };
                        return array.Contains(Cell(16).ToLower());
                    }
                case 9:
                    {
                        string[] array = { "по тер-ти" };
                        return array.Contains(Cell(16).ToLower());
                    }
                case 10:
                    {
                        string[] array = { "в др. суб.", "да", "15" };
                        return array.Contains(Cell(15).ToLower());
                    }
                case 11:
                    {
                        int val = CellInt(19);
                        return (val > 0) && (val <= 10);
                        
                    }
                case 12: return CellInt(19) > 10;
                default:
                    return false;
            }
        }
    }
}

