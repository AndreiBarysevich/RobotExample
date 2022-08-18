using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
//using System.Linq;
using System.Globalization;
using ABBYY;
using ABBYY.FlexiCapture;
using System.IO;
using System.Xml.Serialization;
using System.Threading;


namespace Utils
{
    #region форматирование данных
    public static partial class Date
    {
        /// <summary>
        /// Проверяет дату на валидность и приводит к формату, если это возможно
        /// </summary>
        /// <param name="context">принимает на вход контекст скрипта(IRuleContext)</param>
        public static void ValidateAndFixDate(IRuleContext context)
        {

            if (context == null) throw new ArgumentNullException("context");

            if (string.IsNullOrEmpty(context.Text)) return;

            if (context.Text.ToUpper().IndexOf("Г.") > 0)
                context.Text = context.Text.ToUpper().Replace("Г.", "");

            DateTime date;
            if (!TryParseDate(context.Text, out date))
            {
                context.SetError("Поле не соответствует формату даты.");
                return;
            }

            if (date > DateTime.Today)
            {
                context.SetError("Дата не может быть позднее чем сегодня.");
                return;
            }

            context.Text = date.ToString("d", CultureInfo.GetCultureInfo("ru-RU"));
        }

        public static void OnlyValidateDate(IRuleContext context)
        {
            if (context == null) throw new ArgumentNullException("context");

            if (string.IsNullOrEmpty(context.Text)) return;

            DateTime date;
            if (!TryParseDate(context.Text, out date))
            {
                context.SetError("Поле не соответствует формату даты.");
                return;
            }

            context.Text = date.ToString("d", CultureInfo.GetCultureInfo("ru-RU"));
        }

        /// <summary>
        /// Проверка даты: Дата документа должна быть в формате ДД.ММ.ГГГГ и не позднее текущей даты. 
        /// </summary>
        /// <param name="context"></param>
        public static void IsValidDate(IRuleContext context)
        {
            if (!string.IsNullOrEmpty(context.Text))
            {
                DateTime date;
                string sDate = context.Text;
                if (!DateFormat(ref sDate) || !DateTime.TryParseExact(sDate, "dd.MM.yyyy", new CultureInfo("ru-RU"), DateTimeStyles.None, out date))
                    context.SetError("Поле не соответствует формату даты.");
                else if (date > DateTime.Today)
                    context.SetError("Дата не может быть позднее чем сегодня.");
            }
        }


        public static bool TryParseDate(string s, out DateTime result)
        {
            if (!DateFormat(ref s))
            {
                result = DateTime.MinValue;
                return false;
            }

            return DateTime.TryParseExact(s, "dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU"),
                                          DateTimeStyles.None, out result);
        }

        /// <summary>
        /// форматирование даты
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

        public static string IsValidInn(string INN)
        {
            if (string.IsNullOrEmpty(INN))
            {
                return "Поле ИНН не может быть пустым.";
            }
            if (!Regex.IsMatch(INN, @"^[0-9](\d{11}|\d{9})$"))
            {
                return "ИНН должен состоять из 10 или 12 цифр.";
            };

            int[] key10 = new int[] { 2, 4, 10, 3, 5, 9, 4, 6, 8 };
            int[] key11 = new int[] { 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 };
            int[] key12 = new int[] { 3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8 };

            if (INN.Length == 10)
            {
                int s10 = 0;
                for (int i = 0; i < 9; i++)
                    s10 += key10[i] * int.Parse(INN.Substring(i, 1));
                s10 = (s10 % 11) % 10;
                if (INN.Substring(9, 1) != s10.ToString())
                    return "Контрольное число некорректно, проверьте правильность введенного ИНН.";
            }
            else
            {
                int s11 = 0;
                int s12 = 0;

                for (int i = 0; i < 10; i++)
                    s11 += key11[i] * int.Parse(INN.Substring(i, 1));
                for (int i = 0; i < 11; i++)
                    s12 += key12[i] * int.Parse(INN.Substring(i, 1));

                s11 = (s11 % 11) % 10;
                s12 = (s12 % 11) % 10;

                if (INN.Substring(10, 1) != s11.ToString() || INN.Substring(11, 1) != s12.ToString())
                    return "Контрольное число некорректно, проверьте правильность введенного ИНН.";
            }
            return null;
        }


        private static List<string> flags = new List<string>();

        /// <summary>
        /// Устанавливает CheckSucceeded = false, задает сообщение errorMsg, 
        /// и выставляет HasRuleError = true только для переданных полей.
        /// </summary>
        /// <param name="context">Контекст правил</param>
        /// <param name="errorMsg">Текст ошибки</param>
        /// <param name="focusField">Поле для установки фокуса в ошибке</param>
        /// <param name="errorFields">Поля для включения в ошибку</param>
        public static void SetError(this IRuleContext context, string errorMsg, string focusField = null, params string[] errorFields)
        {
            //добавляет пробел в конце сообщения, через раз, для обхода логики работы флекси (не обновляет поля)
            if (flags.Contains(errorMsg))
            {
                flags.Remove(errorMsg);
                errorMsg += " ";
            }
            else flags.Add(errorMsg);
            context.CheckSucceeded = false;
            context.ErrorMessage = errorMsg;
            if (!string.IsNullOrEmpty(focusField)) context.FocusedField = context.Field(focusField);
            if (errorFields != null && errorFields.Length != 0)
                foreach (IField f in context.Fields)
                    f.HasRuleError = (Array.IndexOf<string>(errorFields, f.Name) > -1);
        }


    }
    #endregion


    #region Очистка полей кроме букв и чисел
    public static partial class Formating
    {
        /// <summary>
        /// Функция очистки строки 
        /// в поле остаются только буквы rus+eng во всех регистрах и цифры
        /// </summary>
        /// <param name="Context"></param>
        /// <returns></returns>//
        public static string LettersAndNumbers(string Context)
        {
            if (Context.Length > 0)
            {
                Regex rgx = new Regex(@"[^а-яА-Яa-zA-Z0-9+\-& \n]");
                Context = rgx.Replace(Context, "");
                Context = Context.Replace($"\n", " ");
                Context = Context.Replace("  ", " ");
                return Context;
            }
            return Context;
        }
        public static string Numbers(string Context)
        {
            if (Context.Length > 0)
            {
                Regex rgx = new Regex("[^0-9]");
                Context = rgx.Replace(Context, "");
                return Context;
            }

            return Context;
        }
    }
    #endregion


    #region Работа с фокусом поля
    public static partial class Region
    {
        /// <summary>
        /// Фокусирует регион одного поля на другое
        /// </summary>
        /// <param name="Context">Контекст</param>
        /// <param name="fieldfrom">Поле с которого нужно взять регион</param>
        /// <param name="fieldto">Поле которому нцужно указать регион</param>
        public static void FocusOn(IRuleContext Context, string fieldfrom, string fieldto)
        {
            Context.Field(fieldto).AddRegion(Context.Field(fieldfrom).Regions[0].Page, Context.Field(fieldfrom).Regions[0].SurroundingRect.ToString());
        }
    }
    #endregion

    #region Форматирование текста в полях (делать переносы строк)
    public static partial class TextFormat
    {
        /// <summary>
        /// Функция нормализации длинны строки (делает переносы строк в длинных полях)
        /// поле считается длинным если >70 
        /// ищет пробел в промежутке от 55 до 70 символа и делает перенос строки
        /// </summary>
        /// <param name="Context"></param>
        /// <returns></returns>//
        public static void FormatText(IRuleContext Context)
        {
            Context.Text = Context.Text.Replace($"-\n", "");
            Context.Text = Context.Text.Replace($"\n", " ");
            Context.Text = Context.Text.Replace("  ", " ");
            if (Context.Text.Length < 71)
            {
            }
            else
            {
                string result = "";
                string text = Context.Text;
                //for (int i = 0; i < Context.Text.Length; i++)
                //{
                while (text.Length > 0)
                {
                    if (text.Length > 70)
                    {
                        int j = text.Substring(50, 20).IndexOf(" ");
                        if (j != -1)
                        {
                            result += text.Substring(0, 50 + j) + $"\r\n";
                            text = text.Substring(50 + j, text.Length - 50 - j);
                        }
                        else
                        {
                            result += text.Substring(0, 70) + $"\r\n";
                            text = text.Substring(70, text.Length - 70);
                        }

                    }
                    else
                    {
                        result += text;
                        text = "";
                    }
                    //if (text.Length < 70 && text.Length > 0)
                    // result += text;
                    // i = result.Length-1;
                }
                Context.Text = result;
            }
        }



    }
    #endregion

    #region for test
    /* public static partial class TextFormat
     {
         /// <summary>
         /// Функция нормализации длинны строки (делает переносы строк в длинных полях)
         /// поле считается длинным если >70 
         /// ищет пробел в промежутке от 55 до 70 символа и делает перенос строки
         /// </summary>
         /// <param name="Context"></param>
         /// <returns></returns>//
         public static string FormatTextY(string Context)
         {
             Context = Context.Replace($"-\n", "");
             Context = Context.Replace($"\n", " ");
             Context = Context.Replace("  ", " ");
             if (Context.Length < 71)
             {
             }
             else
             {
                 string result = "";
                 string text = Context;
                 //for (int i = 0; i < Context.Length; i++)
                 //{
                 while(text.Length>0)
                 {
                     if (text.Length >= 70)
                     {
                         int j = text.Substring(50, 20).IndexOf(" ");
                         if (j != -1)
                         {
                             result += text.Substring(0, 50 + j) + $"\r\n";
                             text = text.Substring(50 + j, text.Length - 50 - j);
                         }
                         else
                         {
                             result += text.Substring(0, 70) + $"\r\n";
                             text = text.Substring(70, text.Length - 70);
                         }
                     }
                     else
                     {
                         result += text;
                         text = "";
                     }
                     //if (text.Length < 70 && text.Length > 0)

                     //i = result.Length;
                 }
                 Context = result;
             }
             return Context;
         }



     }*/
    #endregion
    /*
        try
                    {
                if (text.IndexOf(".")!=-1 || text.IndexOf(",") != -1)
                    Console.WriteLine("{0:0.00}", Convert.ToDecimal(text, culture).ToString());
                else
                    Console.WriteLine("{0:0.00}", Convert.ToDecimal(text, culture).ToString()+".00");
            }
                    catch (FormatException)
                    {
                        Console.WriteLine("FormatException");
                    }
     
     */
    /// <summary>
    /// Метод для работы с датой
    /// </summary>
    public static partial class SummFormat
    {
        /// <summary>
        /// Форматирует суммы
        /// </summary>
        /// <param name="summ"></param>
        /// <returns></returns>
        public static string ValiSumm(string summ)
        {
            if (summ.Length > 0)
            {
                CultureInfo provider;
                //   NumberStyles styles;
                //    styles = NumberStyles.Integer | NumberStyles.AllowDecimalPoint;

                //   if ((styles & NumberStyles.AllowCurrencySymbol) > 0)
                //       provider = new CultureInfo("ru-RU");
                //   else
                provider = CultureInfo.CurrentCulture;

                string k = provider.NumberFormat.NumberDecimalSeparator;


                Regex rgx = new Regex("[^0-9]");
                summ = rgx.Replace(summ, k);

                try
                {

                    bool isDouble = Double.TryParse(summ, out double summD);
                    if (isDouble)
                    {
                        if (summD.ToString("0.00").IndexOf(k) != -1)
                        {
                            if (summD.ToString("0.00").IndexOf(k) == summD.ToString("0.00").Length - 2)
                            {
                                return summD.ToString("0.00").Replace(k, ".") + "0";
                            }
                            else
                            {
                                return summD.ToString("0.00").Replace(k, ".");
                            }
                        }
                        else

                        {
                            return summD.ToString() + "." + "00";
                        }
                    }
                    else
                    {
                        return "FormatException";
                    }
                }
                catch (FormatException)
                {
                    return "FormatException";
                }
            }
            else return summ;

        }

        /// <summary>
        /// Форматирование суммы по красоте
        /// </summary>
        /// <param name="summ"></param>
        /// <returns></returns>
        public static string SummFix(string summ)
        {
            Int32 indexFirs = 0;
            Int32 indexLast;
            bool HaveNum=false;

            if (summ.Length > 0)
            {
                Char[] delays = { '.', ',', '-' };
                string[] exclude = { "-", ",", "..", ",,", ",.", ".,", ".РУБ,", ",РУБ.", ".РУБ.", "РУБ.", "РУБ,", "РУБ", ".." };

                //               summ = Regex.Replace(summ, "[^0-9.,-]", "");
                CultureInfo provider;
                provider = CultureInfo.CurrentCulture;
                string k = provider.NumberFormat.NumberDecimalSeparator;
                summ = summ.Replace(" ", "");
                Regex rgx = new Regex("[^0-9]");
                summ = rgx.Replace(summ, k);

                for (int i = 0; i < exclude.Length; i++)
                {
                    summ = summ.ToUpper().Replace(exclude[i], k);
                }
                summ = summ.Replace(" ", "");
                summ = Regex.Replace(summ, "[^0-9.,-]", "");
                indexLast = summ.Length - 1;
                char[] CH = summ.ToCharArray();
                for (int i = 0; i < summ.Length; i++)
                {
                    if (Char.IsNumber(summ[i]))
                    {
                        HaveNum = true;
                        indexFirs = i;
                        break;
                    }
                }
                for (int i = summ.Length - 1; i >= 0; i--)
                {
                    if (Char.IsNumber(summ[i]))
                    {
                        
                        indexLast = i;
                        break;
                    }
                }
                summ = summ.Substring(indexFirs, indexLast - indexFirs + 1);
                if (HaveNum == true)
                {
                    if (Regex.Matches(summ, "[.,-]").Count > 1)
                    {
                        summ = Regex.Replace(summ, "[.,-]", "");
                        summ = summ.Substring(0, summ.Length - 2) + "." + summ.Substring(summ.Length - 2, 2);
                        SummFix(summ);
                    }
                    else if (summ.IndexOfAny(delays) == summ.Length - 1 || summ.IndexOfAny(delays) == 0)
                    {
                        summ = Regex.Replace(summ, "[.,-]", "");
                        summ = summ + ".00";
                    }
                    else if (Regex.Matches(summ, "[.,-]").Count == 1)
                    {

                        if (summ.IndexOf(".") == summ.Length - 2)
                        {
                            summ += "0";
                        }


                        /* if (summ.Length - summ.IndexOfAny(delays)+1 == 1)
                         {
                             summ = summ.Substring(0, summ.IndexOfAny(delays)) + "." + summ.Substring(summ.IndexOfAny(delays)+1, 1) + "0";
                         }*/

                        if (summ.Substring(summ.Length - 1, 1) == ".")
                            summ = summ.Substring(0, summ.Length - 1);
                    }
                    else
                    { summ = summ + ".00"; }
                }
                else
                { summ = "0.00"; }

            }
            return summ;
        }
    }

    public static partial class GetFields
    {
        /// <summary>
        /// Метод получения значения ячейки таблицы по коду и названию столбца
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="kod"></param>
        /// <param name="col"></param>
        /// <returns>возвращает значение ячейки</returns>
        public static string gettabledata(IDocument Document, IProcessingCallback Processing, string kod, string col)
        {
            string value = "";

            for (int j = 0; j < Document.Sections[0].Children.Count; j++)
            {
                if (Document.Sections[0].Children[j].Type.ToString() == "EFT_Table")
                {
                    string tableName = Document.Sections[0].Children[j].Name;

                    if (Document.Sections[0].HasField(tableName) || Document.Sections[0].Field(tableName).Rows.Count > 0) // проверяем есть ли поле Блок (это наша таблица) и есть ли в таблице строки
                    {
                        Processing.ReportMessage("Начинаем поиск атрибута по коду -  " + kod);

                        for (int i = 0; i < Document.Sections[0].Field(tableName).Rows.Count; i++)
                        {
                            Processing.ReportMessage("Значение ячейки = " + Document.Sections[0].Field(tableName).Cell("Kod", i).Text);
                            if (Document.Sections[0].Field(tableName).Cell("Kod", i).Text.Contains(kod)) //ищем наш код
                            {
                                Processing.ReportMessage("Код найден строка № " + i);
                                value = Document.Sections[0].Field(tableName).Cell(col, i).Value.ToString();
                                break;
                            }
                        }
                        Processing.ReportMessage("Значение найденой ячейки = " + value);
                    }
                }


            }


            return value;
        }

        /// <summary>
        /// Проверка существования ячйки
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="kod"></param>
        /// <returns></returns>
        public static bool hastabledata(IDocument Document, IProcessingCallback Processing, string kod)
        {
            bool value = false;

            for (int j = 0; j < Document.Sections[0].Children.Count; j++)
            {
                if (Document.Sections[0].Children[j].Type.ToString() == "EFT_Table")
                {
                    string tableName = Document.Sections[0].Children[j].Name;

                    if (Document.Sections[0].HasField(tableName) || Document.Sections[0].Field(tableName).Rows.Count > 0) // проверяем есть ли поле Блок (это наша таблица) и есть ли в таблице строки
                    {
                        Processing.ReportMessage("Начинаем поиск атрибута по коду -  " + kod);

                        for (int i = 0; i < Document.Sections[0].Field(tableName).Rows.Count; i++)
                        {
                            Processing.ReportMessage("Значение ячейки = " + Document.Sections[0].Field(tableName).Cell("Kod", i).Text);
                            if (Document.Sections[0].Field(tableName).Cell("Kod", i).Text.Contains(kod)) //ищем наш код
                            {
                                Processing.ReportMessage("Код найден строка № " + i);
                                value = true;
                                break;
                            }
                        }
                        Processing.ReportMessage("Значение найденой ячейки = " + value);
                    }
                }


            }


            return value;
        }

        /// <summary>
        /// метод возвращает значение поля, при его наличии
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="name"></param>
        /// <returns>значение поля или пустое значение</returns>
        public static string GetField(IDocument Document, IProcessingCallback Processing, string name)
        {
            if (Document.Sections[0].HasField(name)) // проверяем есть ли поле
            {
               // Processing.ReportMessage("Document.Sections[0].HasField(name) = " + Document.Sections[0].HasField(name).ToString());
                return Document.Sections[0].Field(name).Text;
            }
            else
            {
               // Processing.ReportMessage("Document.Sections[0].HasField(name) = " + Document.Sections[0].HasField(name).ToString());
                return "";
            }
        }

        /// <summary>
        /// Метод возвращает значение ячейки колонки таблицы по найденому тексту в любой колонке
        /// если нет значения возвращает null
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename">Название таблицы</param>
        /// <param name="searchtext">искомый текст</param>
        /// <param name="column">Колонка в которой искать</param>
        /// <param name="ReturnColumn">колонку которую вернуть</param>
        /// <returns>Text or null</returns>
        public static string GetCellWhere(IDocument Document, IProcessingCallback Processing, string tablename, string searchtext, string column, string ReturnColumn)
        {
           // Processing.ReportMessage("GetCellWhere");
            string value = string.Empty;
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column) || Document.Sections[0].Field(tablename).HasField(ReturnColumn)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    //Processing.ReportMessage("Начинаем поиск '" + searchtext + "' в столбце '" + column + "'");

                    for (int i = 0; i < Document.Sections[0].Field(tablename).Rows.Count; i++)
                    {
                        if (Document.Sections[0].Field(tablename).Cell(column, i).Text.ToLower().Replace(" ", "").Contains(searchtext.ToLower().Replace(" ", ""))) //ищем значение
                        {
                           // Processing.ReportMessage("Значение найдено в строке № " + i);
                            return Document.Sections[0].Field(tablename).Cell(ReturnColumn, i).Value.ToString();
                        }
                        else {
                            //Processing.ReportMessage("Значение не найдено.");
                        }
                    }
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "', '" + ReturnColumn + "' не найдены."); }
            return null;
        }

        /// <summary>
        /// Метод ищет в столбце таблицы значение длинее 5 символов и возвращает индексы строк списком
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename">Название таблицы в которой ищем параграфы</param>
        /// <param name="column">столбец в котором нужно искать заголовки параграфа</param>
        /// <returns>Список (List) индексов строк заголовков параграфов таблицы</returns>
        public static List<int> GetTableParagraph(IDocument Document, IProcessingCallback Processing, string tablename, string column)
        {
            //Processing.ReportMessage("GetTableParagraph");
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    List<int> ParagraphRows = new List<int>() { };
                    for (int i = 0; i < Document.Sections[0].Field(tablename).Rows.Count; i++)
                    {
                        if (Document.Sections[0].Field(tablename).Cell(column, i).Text.Length > 7) //ищем значение длинее 5 символов (считаем, что это Заголовок параграфа таблицы)
                        {
                           // Processing.ReportMessage("Строка № " + i + " является заголовком параграфа");
                            ParagraphRows.Add(i);
                        }
                    }
                    return ParagraphRows;
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "' не найдены."); }
            return null;
        }


        public static List<int> GetTableParagraphByCheckMark(IDocument Document, IProcessingCallback Processing, string tablename, string column)
        {
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    List<int> ParagraphRows = new List<int>() { };
                    for (int i = 0; i < Document.Sections[0].Field(tablename).Rows.Count; i++)
                    {
                        if (Document.Sections[0].Field(tablename).Cell(column, i).Value==true) //ищем значение длинее 5 символов (считаем, что это Заголовок параграфа таблицы)
                        {
                            // Processing.ReportMessage("Строка № " + i + " является заголовком параграфа");
                            ParagraphRows.Add(i);
                        }
                    }
                    return ParagraphRows;
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "' не найдены."); }
            return null;
        }

        /// <summary>
        /// Метод возвращает значение ячейки относительно найденого текста в другой ячейке. 
        /// Поиск выполняется по параграфу таблицы, заданному номерами строк
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename">Название таблицы в которой осуществляется поиск</param>
        /// <param name="searchtext">искомый текст</param>
        /// <param name="column">колонка в которой ищем значение</param>
        /// <param name="ReturnColumn">имя столбца, значение ячейки которого надо вернуть</param>
        /// <param name="start">индекс первой строки параграфа</param>
        /// <param name="end">индекс последней строки параграфа</param>
        /// <returns>значение ячейки</returns>
        public static string GetCellByIndexWhere(IDocument Document, IProcessingCallback Processing, string tablename, string searchtext, string column, string ReturnColumn, int start, int end)
        {
            //Processing.ReportMessage("GetCellByIndexWhere");
           // string value = string.Empty;
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column) || Document.Sections[0].Field(tablename).HasField(ReturnColumn)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    //Processing.ReportMessage("Начинаем поиск '" + searchtext + "' в столбце '" + column + "'");
                   // Processing.ReportMessage("Начальная строка '" + start+ "' последняя строка '" + end + "'");
                    //Processing.ReportMessage("Rows.Count = " + Document.Sections[0].Field(tablename).Rows.Count);

                    if (start == end)
                    {
                       // Processing.ReportMessage("  Rows.Count = " + Document.Sections[0].Field(tablename).Rows.Count);
                        end = Document.Sections[0].Field(tablename).Rows.Count;
                       // Processing.ReportMessage("end = " + end);
                    } //если параграф последний

                    for (int l = start; l < end; l++)
                    {
                       // Processing.ReportMessage("l = " + l);
                     
                        if (Document.Sections[0].Field(tablename).Cell(column, l).Text.ToLower().Replace(" ","").Contains(searchtext.ToLower().Replace(" ", ""))) //ищем значение
                        {
                           // Processing.ReportMessage("Значение найдено в строке № " + l);
                            return Document.Sections[0].Field(tablename).Cell(ReturnColumn, l).Value.ToString();
                        }
                        else {
                           // Processing.ReportMessage("Значение не найдено.");
                        }
                        //if (i == Document.Sections[0].Field(tablename).Rows.Count - 1) { break; };
                        // else { Processing.ReportMessage("Значение не найдено."); }
                    }
                   //  Processing.ReportMessage("Закончили поиск в параграфе и вернули пусто.");
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "', '" + ReturnColumn + "' не найдены."); }
            return null;
        }

        /// <summary>
        /// возвращает значение ячайки по названию столбца и индексу строки
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename">Название таблицы</param>
        /// <param name="column">Название столбца</param>
        /// <param name="index">Индекс строки</param>
        /// <returns></returns>
        public static string GetCellByIndex(IDocument Document, IProcessingCallback Processing, string tablename, string column, int index)
        {
            Processing.ReportMessage("GetCellByIndex");
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column)) // Проверка наличия таблицы и столбцов в ней
            {
               // Processing.ReportMessage("Кол-во строк таблицы: "+ Document.Sections[0].Field(tablename).Rows.Count); 

                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    if (index <= Document.Sections[0].Field(tablename).Rows.Count)
                    {
                        Processing.ReportMessage("Получаем ячейку в таблице: '"+tablename+"'; столбце: "+ column + "; строке: "+ index + "; Значение - "+ Document.Sections[0].Field(tablename).Cell(column, index).Value.ToString());

                        return Document.Sections[0].Field(tablename).Cell(column, index).Value.ToString();
                    }
                    else { Processing.ReportMessage("Строки с индексом '" + index + "' не существует."); }
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "' не найдены."); }
            return null;
        }

        /// <summary>
        /// Метод возвращает список индексов строк, где указаный столбец прошел точное соответсвие регулярному выражению
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename">название таблицы</param>
        /// <param name="column">колонка в которой проверяется соответсвие</param>
        /// <param name="regex">Регулярное выражение</param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public static List<int> GetTableRowsWhereRegex(IDocument Document, IProcessingCallback Processing, string tablename, string column, Regex regex, int start, int end)
        {
           // Processing.ReportMessage("GetTableRowsWhereRegex");
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    List<int> ParagraphRows = new List<int>() { };
                    if (start == end) { end = Document.Sections[0].Field(tablename).Rows.Count; } //если параграф последний
                    for (int i = start; i < end; i++)
                    {
                        if (regex.Match(Document.Sections[0].Field(tablename).Cell(column, i).Text).Value == Document.Sections[0].Field(tablename).Cell(column, i).Text) //проверяем значение ячейки на точное соответсвие регулярному выражению
                        {
                           // Processing.ReportMessage("Строка № '" + i + "' содержит искомое значением");
                            ParagraphRows.Add(i);
                        }
                    }
                    return ParagraphRows;
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "' не найдены."); }
            return null;
        }

        /// <summary>
        /// Возвращает список строк таблицы не содержащих слова из списка
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename"></param>
        /// <param name="column"></param>
        /// <param name="words">список слов</param>
        /// <returns></returns>
        public static List<int> GetTableParagraphWhereCellsNotMachText(IDocument Document, IProcessingCallback Processing, string tablename, string column,List<string> words)
        {
            //Processing.ReportMessage("GetTableParagraph");
            if (Document.Sections[0].HasField(tablename) || Document.Sections[0].Field(tablename).HasField(column)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    List<int> ParagraphRows = new List<int>() { };
                    Processing.ReportMessage("Всего строк в таблице: "+ Document.Sections[0].Field(tablename).Rows.Count);
                    for (int i = 0; i < Document.Sections[0].Field(tablename).Rows.Count; i++)
                    {
                        bool ispararaph = true;
                        foreach (string word in words)
                        {
                            if (Document.Sections[0].Field(tablename).Cell(column, i).Text.ToLower().Replace(@"\r\n", "").Contains(word.ToLower()) || Document.Sections[0].Field(tablename).Cell(column, i).Text.Length<1) //ищем строки не содержащие слова из списка
                            {
                                ispararaph = false;
                            }
                        }
                        if (ispararaph == true)
                        {
                            Processing.ReportMessage("Строка № " + i + " является заголовком параграфа");
                            ParagraphRows.Add(i);
                        }
                    }
                    return ParagraphRows;
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая."); }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '" + column + "' не найдены."); }
            return null;
        }


        /// <summary>
        /// Проверяет наличие слов в столбце
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename"></param>
        /// <param name="column"></param>
        /// <param name="words"></param>
        /// <returns></returns>
        public static bool HasWordsInColumn(IRuleContext Context, string column, List<string> words)
        {
            //Processing.ReportMessage("GetTableParagraph");
            if (Context.HasField(column)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Context.Field(column).Items.Count > 0) //не пустая ли таблица?
                {
                   
                    for (int i = 0; i < Context.Field(column).Items.Count; i++)
                    {
                        foreach (string word in words)
                        {
                            if (Context.Field(column).Items[i].Text.ToLower().Contains(word.ToLower())) //ищем строки содержащие слова из списка
                            {
                                return true;
                            }
                        }
                    }

                    return false;
                }
                else {  }
            }
          
            return false;
        }


        /// <summary>
        /// Проверяет пустая ли таблица
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="tablename">Имя таблицы </param>
        /// <returns></returns>
        public static bool TableIsEmpty(IDocument Document, IProcessingCallback Processing, string tablename)
        {
            if (Document.Sections[0].HasField(tablename)) // Проверка наличия таблицы и столбцов в ней
            {
                if (Document.Sections[0].Field(tablename).Rows.Count > 0) //не пустая ли таблица?
                {
                    foreach (var row in Document.Sections[0].Field(tablename).Rows)
                    {
                        //Document.Sections[0].Field(tablename).Children.Count;
                        for (int i = 0; i < Document.Sections[0].Field("Table1").Rows[0].Children.Count; i++)
                        {
                            Processing.ReportMessage("First row = " + Document.Sections[0].Field("Table1").Rows[0].Children[i].Text);
                        }

                        // foreach (var column in Document.Sections[0].Field(tablename).Rows[row].)
                    }

                    /*
                    if (index <= Document.Sections[0].Field(tablename).Rows.Count)
                    {
                        Processing.ReportMessage("Получаем ячейку в таблице: '" + tablename + "'; столбце: " + column + "; строке: " + index + "; Значение - " + Document.Sections[0].Field(tablename).Cell(column, index).Value.ToString());

                        return Document.Sections[0].Field(tablename).Cell(column, index).Value.ToString();
                    }
                    else { Processing.ReportMessage("Строки с индексом '" + index + "' не существует."); }*/
                }
                else { Processing.ReportMessage("Таблица '" + tablename + "' пустая.");
                    return true;
                }
            }
            else
            { Processing.ReportMessage("Таблица '" + tablename + "' или столбцы '"  + "' не найдены.");
                return true;
            }


            return true;
        }
    }



}