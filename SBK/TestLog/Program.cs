using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.Text.RegularExpressions;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using System.Data;
using Utils;


namespace TestLog
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public static string k()
        {
            CultureInfo provider;
            provider = CultureInfo.CurrentCulture;
            string k = provider.NumberFormat.NumberDecimalSeparator;
            return k;
        }


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

        static void Main()
        {


            var text = Console.ReadLine();
            Console.WriteLine(Regex.Match(text, @"20\d\d").ToString());

            /*
            var date = new DateTime();
            date = DateTime.Now;
            int s  =date.Year - 1;
            Console.WriteLine("year "+ s);
            string text = Console.ReadLine();
            //monthCalendar.Show();*/

            int mounth = 0;

            //DateTime data;
        



            Console.ReadKey();
            Main();
            /*

             // MessageBox.Show();

              //Regex regex = new Regex(@"\d{2,3}");
              Console.WriteLine("введите значение" + Environment.MachineName);
              string s = Console.ReadLine();
              //s = s.Substring(0,s.IndexOf("."));

              Console.WriteLine(s);
             // decimal.TryParse(s.Replace(",.", k()), out decimal a);

             // a = a * 1000;

             // Console.WriteLine(a);
              */
        }







        //Console.WriteLine(Path.GetDirectoryName(Path.GetDirectoryName(s)));
        //Console.ReadKey();


        //logger.Trace("Начало обработки");
        //DateTime date = new DateTime();
        //string text = Console.ReadLine();
        //logger.Debug("В процессе");
        //DateTime.TryParse(text, out date);

        //Console.WriteLine("Результат: " + date);
        // decimal.TryParse(s, out decimal a);
        //Console.WriteLine("decimal- "+a);
        /*
        if (regex.Match(s).Value == s) { Console.WriteLine("Есть матч"); }
        else { Console.WriteLine("Нет матча"); }
        */

      

        //if ("" == null) { Console.WriteLine(" == null"); }
        //else { Console.WriteLine(" != null"); }

    

    /*
    logger.Trace("logger.Trace {0}", "123");
    logger.Debug("logger.Debug");
    logger.Info("logger.Info");
    logger.Warn("logger.Warn");
    logger.Error("logger.Error");
    logger.Fatal("logger.Fatal");*/



        public static void log(string text, string filepath = @"C:\Temp")
        {
            filepath = filepath + @"\" + DateTime.Now.Date.ToString("yyyy.MM.dd")+ @".txt";
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

    }
}
