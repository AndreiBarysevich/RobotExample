using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using ABBYY;
using ABBYY.FlexiCapture;
using System.IO;
using System.Xml.Serialization;
using System.Threading;
using System.Data;
using System.Data.OleDb;

namespace Utils
{
    public class FieldRules
    {
        public static void log(string text, string filepath = @"C:\Temp")
        {
            try
            {
                filepath = filepath + @"\" + DateTime.Now.Date.ToString("yyyy.MM.dd") + @".txt";
                if (!File.Exists(filepath))
                {
                    using (var myfile = File.Create(filepath))
                    {
                        myfile.Close();
                    }
                }

                using (StreamWriter fs = new StreamWriter(filepath, true))
                {
                    fs.WriteLine(DateTime.Now.Date.ToString("G") + "--" + text);
                    fs.Close();
                }
            }
            catch
            {
            }
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


        public static void ModsSuggest(IRuleContext context)
        {
            
            if (!string.IsNullOrEmpty(context.Field("Type").Text) )
            {
                DataTable Data = ConvertExcelToDataTable(@"C:\Temp\Schetchik.xlsx");
                List<string> Mname = new List<string>();
                //log("Text Type = "+ context.Field("Type").Text);
               // log("Text TypeM = " + context.Field("TypeM").Text);
                foreach (DataRow row in Data.Rows)
                {
                   // log("DataRow = SName = " +row["ShortName"].ToString()+ " Modification = "+ row["Modification"].ToString());
                   if (context.Field("Type").Text.ToLower().Contains(row["ShortName"].ToString().ToLower()))
                    {
                       //log(context.Field("Type").Text + " Содержит " + row["ShortName"].ToString() + " - Добавляем! значение модимфикации = "+ row["Modification"].ToString()+ "");
                        Mname.Add(""+row["Modification"].ToString());
                    }
                }
                Mname.Add("");
               // log("Список из модификайций: ");
                foreach (string item in Mname)
                {
                   // log(item);
                    context.Field("TypeM").Suggest(item);
                }
                
            }
        }


        public static void NamesSuggest(IRuleContext context)
        {
            
            DataTable Data = ConvertExcelToDataTable(@"C:\Temp\Schetchik.xlsx");
           
            List<string> Sname = new List<string>();
            foreach (DataRow row in Data.Rows)
            {
                Sname.Add(row["ShortName"].ToString());
            }
           
            List<string> name = new List<string>();
            name = GetNamesContains(Sname, context.Field("Type").Text);
            foreach (string item in name)
            {
                context.Suggest(item);
            }
            if (!context.IsVerified)
            {
                foreach (string item in name)
                {
                    if (item.ToLower().Contains(context.Text))
                    {
                        context.Text = item;
                        break;
                    }
                    else if (GetDistanceCore(item.ToLower(), context.Text.ToLower()) < 0.3)
                    {
                        context.Text = item;
                        break;
                    }
                }
            }
        }

        public static List<string> GetNamesContains(List<string> Names, string text)
        {
            List<string> name = new List<string>();
           
            foreach (string texts in Names)
            {
               
                if (text.Length > 3)
                {
                    
                    if (texts.ToLower().Contains(text.ToLower()))
                    {
                      
                        name.Add(texts);
                    }
                    else if (GetDistanceCore(texts.ToLower(), text.ToLower()) < 0.3)
                    {
                     
                        name.Add(texts);
                    }
                    else { }
                }

            }

            return name;
        }


        public static void SuggestPeriod(IRuleContext context)
        {
            context.Suggest("январь");
            context.Suggest("февраль");
            context.Suggest("март");
            context.Suggest("апрель");
            context.Suggest("май");
            context.Suggest("июнь");
            context.Suggest("июль");
            context.Suggest("август");
            context.Suggest("сентябрь");
            context.Suggest("октябрь");
            context.Suggest("ноябрь");
            context.Suggest("декабрь");
            
            string text = context.Text;
            if (!Regex.IsMatch(text, @"январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь", RegexOptions.IgnoreCase))
            {
                context.Text = string.Empty;
                int month = 0;
                if (string.IsNullOrEmpty(text))
                {
                    context.SetError("Укажите Период");
                    // month = string.Empty;
                    return;
                } //если значение пусто
                else if (DateFormat(ref text) != false) //если значение в виде даты
                {
                    // Console.WriteLine("Date " + text);
                    DateTime.TryParse(text, out DateTime date);
                    month = date.Month;
                }
                else if (Regex.IsMatch(text, @"([1234]( |)квартал)|(квартал( |)[1234])", RegexOptions.IgnoreCase)) //если это квартал
                {
                    if (Regex.IsMatch(text, @"(1( |)Квартал)|(Квартал( |)1)", RegexOptions.IgnoreCase)) { month = 3; }
                    else if (Regex.IsMatch(text, @"(2( |)Квартал)|(Квартал( |)2)", RegexOptions.IgnoreCase)) { month = 6; }
                    else if (Regex.IsMatch(text, @"(3( |)Квартал)|(Квартал( |)3)", RegexOptions.IgnoreCase)) { month = 9; }
                    else if (Regex.IsMatch(text, @"(4( |)Квартал)|(Квартал( |)4)", RegexOptions.IgnoreCase)) { month = 12; }
                }
                else if (Regex.Matches(text, @"([01][0987654321])|[123456789]", RegexOptions.IgnoreCase).Count > 0)
                {
                    int[] monthindex = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
                    for (int mi = 0; mi < monthindex.Length; mi++)
                    {
                        if (text.Contains(monthindex[mi].ToString()))
                        {
                            month = monthindex[mi];
                            break;
                        }
                    }
                }
                else if (Regex.Matches(text, @"январ[ья]|янв|феврал[ья]|фев|март[а]|мар|апрел[ья]|апр|ма[йя]|июн[ья]{0,1}|июл[ья]{0,1}|август[а]{0,1}|авг|сентябр[ья]|сен[т]{0,1}|октябр[ья]|окт|ноябр[ья]|ноя|декабр[ья]{0,1}|дек").Count > 0)
                {
                    string[] monthPatterns1 = { "январ[ья]|янв", "феврал[ья]|фев", "март[а]|мар", "апрел[ья]|апр", "ма[йя]", "июн[ья]{0,1}", "июл[ья]{0,1}", "август[а]{0,1}|авг", "сентябр[ья]|сен[т]{0,1}", "октябр[ья]|окт", "ноябр[ья]|ноя", "декабр[ья]{0,1}|дек" };
                    string[] monthReplacement1 = { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" };
                    //bool done = false;
                    for (int monthIndex = 0; monthIndex < monthPatterns1.Length; monthIndex++)
                    {
                        if (!Regex.IsMatch(text, monthReplacement1[monthIndex])) continue;
                        {
                            Int32.TryParse(monthReplacement1[monthIndex], out month);
                            //done = true;
                            //Console.WriteLine(text);
                            break;
                        }
                    }

                }
                else
                {
                    context.SetError("Период содержит некорректное значение");
                }

                if (month == 0) { context.Text = string.Empty; }
                else if (month == 1) { context.Text = "январь"; }
                else if (month == 2) { context.Text = "февраль"; }
                else if (month == 3) { context.Text = "март"; }
                else if (month == 4) { context.Text = "апрель"; }
                else if (month == 5) { context.Text = "май"; }
                else if (month == 6) { context.Text = "июнь"; }
                else if (month == 7) { context.Text = "июль"; }
                else if (month == 8) { context.Text = "август"; }
                else if (month == 9) { context.Text = "сентябрь"; }
                else if (month == 10) { context.Text = "октябрь"; }
                else if (month == 11) { context.Text = "ноябрь"; }
                else if (month == 12 ){ context.Text = "декабрь"; }
                context.IsVerified = true;
            }
            //else
            //{
            //    if (string.IsNullOrEmpty(context.Text))
            //    {
            //        context.SetError("Заполните значение");
            //    }
            //    else if (!Regex.IsMatch(text, @"январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь", RegexOptions.IgnoreCase))
            //    {  }
            //    else
            //        context.IsVerified = true;
            //}
        }


        public static void SuggestYear(IRuleContext context)
        {
            var date = new DateTime();
            date = DateTime.Now;
            int year = date.Year;
            string text = context.Text;

            context.Suggest(year.ToString());
            context.Suggest((year - 1).ToString());
            context.Suggest((year - 2).ToString());
            context.Suggest((year - 3).ToString());
            context.Suggest((year - 4).ToString());
            context.Suggest((year - 5).ToString());
           // context.Suggest((year + 1).ToString());

            if (!context.IsVerified && Regex.IsMatch(context.Text, @"20\d\d"))
            {
                if (string.IsNullOrEmpty(text))
                {

                    context.SetError("Укажите год");
                }
                else
                {
                    if (Regex.IsMatch(context.Text, @"20\d\d"))
                    {
                        text = Regex.Match(context.Text, @"20\d\d").ToString();

                    }
                    else
                    {
                        text = string.Empty;
                    }

                    if (string.IsNullOrEmpty(text))
                    {
                        context.Text = string.Empty;
                        context.SetError("Укажите год");
                    }
                    else context.Text = text;
                }

            }
            else if (!Regex.IsMatch(context.Text, @"20\d\d"))
            {
                context.SetError("Год содержит некорректные данные");
            }

        }


        public static void DateFromPeriod(IRuleContext context)
        {
            if (context.Field("PeriodMonth").IsVerified && context.Field("PeriodYear").IsVerified)
            {
                string mnth = string.Empty;
                string year = context.Field("PeriodYear").Text;
                
                if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"январь", RegexOptions.IgnoreCase)) { mnth = "01"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"февраль", RegexOptions.IgnoreCase)) { mnth = "02"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"март", RegexOptions.IgnoreCase)) { mnth = "03"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"апрель", RegexOptions.IgnoreCase)) { mnth = "04"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"май", RegexOptions.IgnoreCase)) { mnth = "05"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"июнь", RegexOptions.IgnoreCase)) { mnth = "06"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"июль", RegexOptions.IgnoreCase)) { mnth = "07"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"август", RegexOptions.IgnoreCase)) { mnth = "08"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"сентябрь", RegexOptions.IgnoreCase)) { mnth = "09"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"октябрь", RegexOptions.IgnoreCase)) { mnth = "10"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"ноябрь", RegexOptions.IgnoreCase)) { mnth = "11"; }
                else if (Regex.IsMatch(context.Field("PeriodMonth").Text, @"декабрь", RegexOptions.IgnoreCase)) { mnth = "12"; }

                context.Text = "30." + mnth + "." + year;
                Date.ValidateAndFixDate(context);
            }
        }

        /// <summary>
        /// преобразуем дату
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static Boolean DateFormat(ref string date)
        {
            if (string.IsNullOrEmpty(date)) return false;
            //Если дата соответствует формату ДД.ММ.ГГГГ, то не производим дальнейших преобразований
            if (Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[.](0[1-9]|1[012])[.]([0-9]{4})$")) return true;

            //Удаляем пробелы и кавычки
            date = Regex.Replace(date, "[ \"]+", "");
            //Заменяем возможные разделители на точки
            date = Regex.Replace(date, "[/:,-]+", ".");
            //Заменяем многоточия на точку
            date = Regex.Replace(date, "[.]{2,}", ".");
            //Удаляем точки вначале и в конце даты
            if (date[0] == '.') date = date.Remove(0, 1);
            if (date[date.Length - 1] == '.') date = date.Remove(date.Length - 1, 1);

            //Добавляем точки где их не хватает (Стандартная дата)
            if (Regex.IsMatch(date, @"^(\d{6}|\d{8})$")) date = date.Insert(4, ".").Insert(2, ".");
            if (Regex.IsMatch(date, @"^(\d{2}[.](\d{4}|\d{6}))$")) date = date.Insert(5, ".");
            if (Regex.IsMatch(date, @"^(\d{4}[.](\d{2}|\d{4}))$")) date = date.Insert(2, ".");



            if (Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[.](0[1-9]|1[012])[.]([0-9]{4})$")) return true;


            date = date.ToLower();

            //Добавляем точки для числовой даты (Не стандартная дата)
            if (Regex.IsMatch(date, @"^([1-9](1[012])(\d{2}|\d{4}))$")) date = date.Insert(3, ".").Insert(1, ".");
            if (Regex.IsMatch(date, @"^([1-9][.](1[012])(\d{2}|\d{4}))$")) date = date.Insert(4, ".");
            if (Regex.IsMatch(date, @"^([1-9](1[012])[.](\d{2}|\d{4}))$")) date = date.Insert(1, ".");
            if (Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[1-9](\d{2}|\d{4})$"))
                date = date.Insert(3, ".").Insert(2, ".");
            if (Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[.][1-9](\d{2}|\d{4})$")) date = date.Insert(4, ".");
            if (Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[1-9][.](\d{2}|\d{4})$")) date = date.Insert(2, ".");
            //Добавляем точки где их не хватает для месяца записанного текстом (Не стандартная дата)
            if (Regex.IsMatch(date, @"^(\d{2}[а-я]{3,8}\d{4})$"))
                date = date.Insert(date.Length - 4, ".").Insert(2, ".");
            if (Regex.IsMatch(date, @"^(\d{2}[а-я]{3,8}\d{2})$"))
                date = date.Insert(date.Length - 2, ".").Insert(2, ".");
            if (Regex.IsMatch(date, @"^([1-9][а-я]{3,8}\d{4})$"))
                date = date.Insert(date.Length - 4, ".").Insert(1, ".");
            if (Regex.IsMatch(date, @"^([1-9][а-я]{3,8}\d{2})$"))
                date = date.Insert(date.Length - 2, ".").Insert(1, ".");
            if (Regex.IsMatch(date, @"^(\d{2}[.][а-я]{3,8}\d{4})$")) date = date.Insert(date.Length - 4, ".");
            if (Regex.IsMatch(date, @"^(\d{2}[.][а-я]{3,8}\d{2})$")) date = date.Insert(date.Length - 2, ".");
            if (Regex.IsMatch(date, @"^([1-9][.][а-я]{3,8}\d{4})$")) date = date.Insert(date.Length - 4, ".");
            if (Regex.IsMatch(date, @"^([1-9][.][а-я]{3,8}\d{2})$")) date = date.Insert(date.Length - 2, ".");
            if (Regex.IsMatch(date, @"^(\d{2}[а-я]{3,8}[.](\d{2}|\d{4}))$")) date = date.Insert(2, ".");
            if (Regex.IsMatch(date, @"^([1-9][а-я]{3,8}[.](\d{4}|\d{2}))$")) date = date.Insert(1, ".");

            //Добавляем число 0 в день\месяц если он отсутствует
            if (Regex.IsMatch(date, @"^(\d{1}[.](\d{1,2}|\D{3,8})[.](\d{2}|\d{4}))$")) date = date.Insert(0, "0");
            if (Regex.IsMatch(date, @"^(\d{2}[.]\d{1}[.](\d{2}|\d{4}))$")) date = date.Insert(3, "0");

            ////Заменяем число 9 на 0 в дне\месяце
            //if (Regex.IsMatch(date, @"^([9]\d{1}[.](\d{2}|\D{3,8})[.](\d{2}|\d{4}))$"))
            //    date = date.Remove(0, 1).Insert(0, "0");
            //if (Regex.IsMatch(date, @"^(\d{2}[.][9]\d{1}[.](\d{2}|\d{4}))$")) date = date.Remove(3, 1).Insert(3, "0");

            //Добавляем "20" в год если он в формате YY, для преобразования в формат YYYY
            if (Regex.IsMatch(date, @"^(\d{2}[.](\d{2}|\D{3,8})[.]\d{2})$")) date = date.Insert(date.Length - 2, "20");

            if (Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[.](0[1-9]|1[012])[.]([0-9]{4})$")) return true;

            //для даты написанной в другой последовательности
            if (Regex.IsMatch(date, @"^([0-9]{4})[.](0[1-9]|1[012])[.](\d|0[1-9]|1\d|2\d|3[01])$"))
            {
                var n = date.Substring(8);
                if (n.Length == 1) n = "0" + n;
                date = string.Format("{0}.{1}.{2}", n, date.Substring(5, 2), date.Substring(0, 4));
                return true;
            }


            //Обрабатываем дату с текстовой записью месяца
            const string datePattern1 = @"^(\d{2}[.](";
            const string datePattern2 = @")[.]\d{4})$";
            string[] monthPatterns =
            {
                "январ[ья]|янв", "феврал[ья]|фев", "март[а]|мар", "апрел[ья]|апр", "ма[йя]",
                "июн[ья]{0,1}", "июл[ья]{0,1}", "август[а]{0,1}|авг", "сентябр[ья]|сен[т]{0,1}", "октябр[ья]|окт",
                "ноябр[ья]|ноя",
                "декабр[ья]{0,1}|дек"
            };
            string[] monthReplacement = { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" };
            for (int monthIndex = 0; monthIndex < monthPatterns.Length; monthIndex++)
            {
                if (!Regex.IsMatch(date, datePattern1 + monthPatterns[monthIndex] + datePattern2)) continue;
                date = Regex.Replace(date, monthPatterns[monthIndex], monthReplacement[monthIndex]);
            }

            return Regex.IsMatch(date, @"^(0[1-9]|1\d|2\d|3[01])[.](0[1-9]|1[012])[.]([0-9]{4})$");
        }


        /// <summary>
        /// пробуем получить дату
        /// </summary>
        /// <param name="s"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public static bool TryParseDate(ref string s, out DateTime result)
        {
            if (!DateFormat(ref s))
            {
                result = DateTime.MinValue;
                return false;
            }

            return DateTime.TryParseExact(s, "dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU"),
                                          DateTimeStyles.None, out result);
        }



    }
}
