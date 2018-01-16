using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter
{
    class Item
    {
        public int[] c = { 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        public String Solution = String.Empty;
        public String Issue = String.Empty;
        public String Outgo = String.Empty;
        public String OutgoT = String.Empty;

        public string id { get; set; }
        public String trace = String.Empty;

        public int Count { get { return c[0]; } set { c[0] = value; } }

        public Boolean Is2()
        {
            return Solution.IndexOf("СЕ", StringComparison.Ordinal) >= 0;
        }
        public Boolean Is3()
        {
            return Solution.IndexOf("УД", StringComparison.Ordinal) >= 0;
        }
        public Boolean Is4()
        {
            return Solution == "отказано";
        }

        public Boolean Is5()
        {
            string[] array = { "1", "2" };
            return array.Contains(Issue);
        }
        public Boolean Is6()
        {

            return Issue == "3";
        }
        public Boolean Is7()
        {
            return Issue == "4";
        }
        public Boolean Is8()
        {
            string[] array = { "по под-ти", "по тер-ти" };
            return array.Contains(Solution);
        }
        public Boolean Is9()
        {
            string[] array = { "по под-ти" };
            return array.Contains(Solution);
        }
        public Boolean Is10()
        {
            string[] array = { "по тер-ти" };
            return array.Contains(Solution);
        }
        public Boolean Is11()
        {
            string[] array = { "в др. суб.", "да", "15" };
            return array.Contains(Outgo);
        }
        public Boolean Is12()
        {
            string[] array = { "3", "10" };
            return array.Contains(OutgoT);
        }
        public Boolean Is13()
        {
            string[] array = { "19", "30" };
            return array.Contains(OutgoT);
        }
    }
}

