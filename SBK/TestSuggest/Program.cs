using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
//using ABBYY;
//using ABBYY.FlexiCapture;
using System.IO;
using System.Xml.Serialization;
using System.Threading;

namespace TestSuggest
{
    class Program
    {
        static void Main(string[] args)
        {
           DataTable Dtable = ConvertExcelToDataTable(@"C:\Temp\Schetchik.xlsx");


            foreach (DataRow row in Dtable.Rows)
            {
                Console.WriteLine(row["Modification"]);

            }
            Console.ReadKey();

        }

        public static bool searchword(string word, string text)
        {

            text = text.ToLower().Replace(" ", "");
            word = word.ToLower().Replace(" ", "");

            // double minIndex = 1;

            //List<string,int> results = new List<string,int> ();
            Dictionary<string, int> results = new Dictionary<string, int>();

            //Console.WriteLine();
            while (text.Length >= word.Length)
            {

                //Console.WriteLine("Ищем в: "+text.Substring(0,word.Length));
                if (GetDistanceCore(text.Substring(0, word.Length), word) < 0.3)
                {

                    Console.WriteLine("Расстояние левенштейна: " + GetDistanceCore(text.Substring(0, word.Length), word));
                    return true;

                }
                // Console.WriteLine(minIndex);
                text = text.Substring(1);
            }
            return false;
        }
        /// <summary>
        /// Реализует алгоритм определня коэфицента схожести не пустых строк.
        /// </summary>
        /// <param name="word1">Первая строка.</param>
        /// <param name="word2">Вторая строка.</param>
        /// <returns>Коэфицентс схожести.</returns>
        public static double GetDistanceCore(string word1, string word2)
        {
            var score = new int[word1.Length + 2, word2.Length + 2];

            var infinityScore = word1.Length + word2.Length;
            score[0, 0] = infinityScore;
            for (var i = 0; i <= word1.Length; i++)
            {
                score[i + 1, 1] = i;
                score[i + 1, 0] = infinityScore;
            }

            for (var j = 0; j <= word2.Length; j++)
            {
                score[1, j + 1] = j;
                score[0, j + 1] = infinityScore;
            }

            var sd = new SortedDictionary<char, int>();
            foreach (var letter in (word1 + word2))
            {
                if (!sd.ContainsKey(letter))
                    sd.Add(letter, 0);
            }

            for (var i = 1; i <= word1.Length; i++)
            {
                var DB = 0;
                for (var j = 1; j <= word2.Length; j++)
                {
                    var i1 = sd[word2[j - 1]];
                    var j1 = DB;

                    if (word1[i - 1] == word2[j - 1])
                    {
                        score[i + 1, j + 1] = score[i, j];
                        DB = j;
                    }
                    else
                    {
                        score[i + 1, j + 1] = Math.Min(Math.Min(score[i, j], score[i + 1, j]), score[i, j + 1]) + 1;
                    }

                    score[i + 1, j + 1] = Math.Min(score[i + 1, j + 1], score[i1, j1] + (i - i1 - 1) + 1 + (j - j1 - 1));
                }

                sd[word1[i - 1]] = i;
            }

            double maxLength = Math.Max(word1.Length, word2.Length);

            return score[word1.Length + 1, word2.Length + 1] / maxLength;
        }

        public static DataTable ConvertExcelToDataTable(string FileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }

       // public static void NamesSuggest(IRuleContext context)
      

        //public static List<string> GetNamesContains(List<string> Names, string text)
        //{
        //    List<string> name = new List<string>();

        //    foreach (string texts in Names)
        //    {
        //        if (text.Length > 3)
        //        {
        //            if (text.Contains(texts))
        //            {
        //                name.Add(texts);
        //            }
        //        }

        //    }

        //    return name;
        //}

    }
}
