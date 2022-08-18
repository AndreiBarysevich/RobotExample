using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using ABBYY;
using ABBYY.FlexiCapture;
using System.IO;
using System.Xml.Serialization;
using System.Threading;
using System.Text.RegularExpressions;

namespace Utils
{


    public static class ExportXML
    {
        public static ПериодКварталТип PeriodQuarterType(string date)
        {

            if (!string.IsNullOrEmpty(date) && DateTime.TryParse(date, out var dateValue1))
            {
                if (dateValue1.Month == 03 || dateValue1.Month == 04) { return ПериодКварталТип.Item21; }
                else if (dateValue1.Month == 06 || dateValue1.Month == 07) { return ПериодКварталТип.Item31; }
                else if (dateValue1.Month == 09 || dateValue1.Month == 10) { return ПериодКварталТип.Item33; }
                else if (dateValue1.Month == 12 || dateValue1.Month == 01) { return ПериодКварталТип.Item34; } else { return ПериодКварталТип.Item21; }
            }
            else
            {
                if (DateTime.Now.Month > 10) { return ПериодКварталТип.Item34; }
                else if (DateTime.Now.Month > 07) { return ПериодКварталТип.Item33; }
                else if (DateTime.Now.Month > 04) { return ПериодКварталТип.Item31; }
                else if (DateTime.Now.Month >= 01) { return ПериодКварталТип.Item21; } else { return ПериодКварталТип.Item21; }
            }

        }
        public static ПериодМесяцТип PeriodMonthType(string date)
        {

            if (!string.IsNullOrEmpty(date) && DateTime.TryParse(date, out var dateValue1))
            {
                if (dateValue1.Month == 01) { return ПериодМесяцТип.Item35; }
                else if (dateValue1.Month == 02) { return ПериодМесяцТип.Item36; }
                else if (dateValue1.Month == 03) { return ПериодМесяцТип.Item21; }
                else if (dateValue1.Month == 04) { return ПериодМесяцТип.Item38; }
                else if (dateValue1.Month == 05) { return ПериодМесяцТип.Item39; }
                else if (dateValue1.Month == 06) { return ПериодМесяцТип.Item31; }
                else if (dateValue1.Month == 07) { return ПериодМесяцТип.Item41; }
                else if (dateValue1.Month == 08) { return ПериодМесяцТип.Item42; }
                else if (dateValue1.Month == 09) { return ПериодМесяцТип.Item33; }
                else if (dateValue1.Month == 10) { return ПериодМесяцТип.Item44; }
                else if (dateValue1.Month == 11) { return ПериодМесяцТип.Item45; }
                else if (dateValue1.Month == 12) { return ПериодМесяцТип.Item34; } else { return ПериодМесяцТип.Item35; }
            }
            else
            {
                return ПериодМесяцТип.Item35;
            }

        }

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

        /// <summary>
        /// Возвращает разделитель текущей машины
        /// </summary>
        /// <returns></returns>
        public static string k()
        {
            CultureInfo provider;
            provider = CultureInfo.CurrentCulture;
            string k = provider.NumberFormat.NumberDecimalSeparator;
            return k;
        }
        // private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Проверяет наличие индекса в списке индексов
        /// </summary>
        /// <param name="data"></param>
        /// <param name="plist"></param>
        /// <returns></returns>
        public static bool hasitem(int data, List<int> plist)
        {
            foreach (int item in plist)
            {
                if (item == data)
                { return true; }
            }
            return false;
        }

        /// <summary>
        /// Метод экспорта документов по XSD схеме
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="iso"></param>
        /// <param name="ExpPath"></param>
        /// <param name="NeedOrigins"></param>
        public static void Parsing(IDocument Document, IProcessingCallback Processing, IExportImageSavingOptions iso, string ExpPath = "", bool NeedOrigins = false)
        {
            Processing.ReportMessage("Начинаем экспорт данных ");
            #region Определяем папку экспорта
            if (Document.Batch.Properties.Has("Export") && !string.IsNullOrEmpty(Document.Batch.Properties.Get("Export")))
            {
                try
                {
                    Processing.ReportMessage("Path.GetDirectoryName(Document.Batch.Properties.Get(Export))" + Path.GetDirectoryName(Document.Batch.Properties.Get("Export")));
                    ExpPath = Path.GetDirectoryName(Document.Batch.Properties.Get("Export") + @"\" + Document.Batch.Name + @"\");//  + "_" +Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource)) +
                    Processing.ReportMessage("Путь экспорта из рег. параметра: " + ExpPath);
                }
                catch
                {
                    Processing.ReportWarning("В Рег.параметре пакета указан не корректный путь экспорта. Экспорт производится относительно папки импорта");
                    ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name +  @"\";//"_"  ++ Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))
                }
            }
            else
            { ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name +  @"\"; }//"_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource)) +
            Processing.ReportMessage("Документ будет экспортироваться в " + ExpPath);

            

            if (!Directory.Exists(ExpPath))
                Directory.CreateDirectory(ExpPath);


            if (NeedOrigins)
            {
                ExportOriginals(Document, Processing, ExpPath);
                Processing.ReportMessage("Оригиналы выгружены");

            }
            #endregion Определяем папку экспорта


            ExpPath = ExpPath + @"\" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource)) + "_" + Document.DefinitionName +"_"+ GetFields.GetField(Document, Processing, "Date") + "_" + Document.Id + ".xml"; //добавляем название файла


            // Уберём расширение из имени файла.
            var serializer = new XmlSerializer(typeof(Файл));
            var stream = new StreamWriter(ExpPath);
            #region Файл
            var Файл = new Файл();
            Файл.ВерсПрог = "ABI:VERSION_0.2";
            Файл.ИдФайл = "1-0";
            Файл.ВерсФорм = ФайлВерсФорм.Item508ABI;
            try
            {
                #region Документы

                #region ШАПКА_xml
                var Документ = new ФайлДокумент();
                //  if (GetFields.GetField(Document, Processing, "NomKorr") != "") { Документ.НомКорр = GetFields.GetField(Document, Processing, "NomKorr"); } else { Документ.НомКорр = "0"; }

                #region СвНп
                ФайлДокументСвНП СвНп = new ФайлДокументСвНП();

                if (GetFields.GetField(Document, Processing, "INN").Length == 10)
                {
                    ФайлДокументСвНПНПЮЛ НПЮЛ = new ФайлДокументСвНПНПЮЛ();
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Organization"))) { НПЮЛ.НаимОрг = GetFields.GetField(Document, Processing, "Organization"); } else { НПЮЛ.НаимОрг = "-"; }
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Organization"))) { НПЮЛ.НаимОргКраткое = GetFields.GetField(Document, Processing, "Organization"); } else { НПЮЛ.НаимОргКраткое = "-"; }
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "INN")) && GetFields.GetField(Document, Processing, "INN").Count() != 12) { НПЮЛ.ИННЮЛ = GetFields.GetField(Document, Processing, "INN"); } else { НПЮЛ.ИННЮЛ = "2222222223"; }
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "KPP"))) { НПЮЛ.КПП = GetFields.GetField(Document, Processing, "KPP"); } //else { НПЮЛ.КПП = "000000000"; }
                    СвНп.Item = НПЮЛ;
                }
                else if (GetFields.GetField(Document, Processing, "INN").Length == 12)
                {
                    ФайлДокументСвНПНПФЛ НПФЛ = new ФайлДокументСвНПНПФЛ();
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "INN")) && GetFields.GetField(Document, Processing, "INN").Count() != 10) { НПФЛ.ИННФЛ = GetFields.GetField(Document, Processing, "INN"); } else { НПФЛ.ИННФЛ = "2222222223"; }
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Organization"))) { НПФЛ.ФИО.Имя = GetFields.GetField(Document, Processing, "Organization"); } else { НПФЛ.ФИО.Имя = "-"; }
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Organization"))) { НПФЛ.ФИО.Фамилия = GetFields.GetField(Document, Processing, "Organization"); } else { НПФЛ.ФИО.Фамилия = "-"; }
                    СвНп.Item = НПФЛ;
                }
                else if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "INN")))
                {
                    ФайлДокументСвНПНПЮЛ НПЮЛ = new ФайлДокументСвНПНПЮЛ();
                    НПЮЛ.НаимОрг = "-";
                    НПЮЛ.НаимОргКраткое = "-";
                    НПЮЛ.ИННЮЛ = "2222222223";
                    //НПЮЛ.КПП = "000000000";
                    СвНп.Item = НПЮЛ;
                }

                #endregion СвНп

                #region Подписант
                var Подписант = new ФайлДокументПодписант();

                var ПрПодп = new ФайлДокументПодписантПрПодп();
                if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "PrPodp")))
                {
                    if (GetFields.GetField(Document, Processing, "PrPodp") == "1")
                    {
                        ПрПодп = ФайлДокументПодписантПрПодп.Item1;
                    }
                    else
                    {
                        if (GetFields.GetField(Document, Processing, "PrPodp") == "2")
                        {
                            ПрПодп = ФайлДокументПодписантПрПодп.Item2;
                        }

                    }
                }
                var ФИО = new ФИОТип();
                if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Fio"))) { ФИО.Имя = (GetFields.GetField(Document, Processing, "Fio")); } else { ФИО.Имя = "-"; }
                if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Fio"))) { ФИО.Фамилия = (GetFields.GetField(Document, Processing, "Fio")); } else { ФИО.Фамилия = "-"; }
                if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Fio"))) { ФИО.Отчество = (GetFields.GetField(Document, Processing, "Fio")); } else { ФИО.Отчество = "-"; }

                // var СвПред = new ФайлДокументПодписантСвПред();
                //СвПред.НаимДок = "";

                Подписант.ФИО = ФИО;
                //  Подписант.СвПред = СвПред;
                Подписант.ПрПодп = ПрПодп;
                #endregion Подписант
                #endregion ШАПКА

                #region Документ_Баланс

                if (Document.DefinitionName == "3_37_2_Бухгалтерский баланс" ||
                    Document.DefinitionName == "3_37_3_Бухгалтерский баланс ОКУД 0710099" || Document.DefinitionName ==
                    "3_37_4_Упрощенная бухгалтерская (финансовая) отчетность")
                {
                    var Баланс = new ФайлДокументБаланс();
                    var БалансОКУД = new ФайлДокументБалансОКУД();
                    var БалансОКЕИ = new ФайлДокументБалансОКЕИ();
                    var БалансПериод = new ПериодКварталТип();
                    БалансПериод = PeriodQuarterType(GetFields.GetField(Document, Processing, "Date"));

                    #region Актив
                    if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Period3"))))
                    {
                        var Aктив = new ФайлДокументБалансАктив();
                        if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Poya"))) Aктив.Пояснения = GetFields.gettabledata(Document, Processing, "1600", "Poya");
                        if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Period1"))) Aктив.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1600", "Period1")).ToString();
                        if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Period2"))) Aктив.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1600", "Period2")).ToString();
                        if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1600", "Period3"))) Aктив.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1600", "Period3")).ToString();

                        #region ВнеОбА
                        if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Period3"))) || GetFields.hastabledata(Document, Processing, "1110") || GetFields.hastabledata(Document, Processing, "1120") || GetFields.hastabledata(Document, Processing, "1130") || GetFields.hastabledata(Document, Processing, "1140") || GetFields.hastabledata(Document, Processing, "1150") || GetFields.hastabledata(Document, Processing, "1160") || GetFields.hastabledata(Document, Processing, "1170") || GetFields.hastabledata(Document, Processing, "1180") || GetFields.hastabledata(Document, Processing, "1190"))
                        {
                            var ВнеОбА = new ФайлДокументБалансАктивВнеОбА();
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Poya"))) { ВнеОбА.Пояснения = GetFields.gettabledata(Document, Processing, "1100", "Poya"); }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Period1"))) { ВнеОбА.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1100", "Period1")).ToString(); } else { ВнеОбА.СумОтч = "0"; }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Period2"))) { ВнеОбА.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1100", "Period2")).ToString(); } else { ВнеОбА.СумПрдщ = "0"; }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1100", "Period3"))) { ВнеОбА.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1100", "Period3")).ToString(); } else { ВнеОбА.СумПрдшв = "0"; }

                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1110", "Poya"))) || (GetFields.gettabledata(Document, Processing, "1110", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1110", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1110", "Period3") != ""))
                            {
                                ОПП_ВПТип ВнеОбАНематАкт = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1110", "Poya"))) ВнеОбАНематАкт.Пояснения = GetFields.gettabledata(Document, Processing, "1110", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1110", "Period1"))) ВнеОбАНематАкт.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1110", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1110", "Period2"))) ВнеОбАНематАкт.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1110", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1110", "Period3"))) ВнеОбАНематАкт.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1110", "Period3")).ToString();
                                ВнеОбА.НематАкт = ВнеОбАНематАкт;
                            }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Poya")) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАРезИсслед = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Poya"))) ВнеОбАРезИсслед.Пояснения = GetFields.gettabledata(Document, Processing, "1120", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Period1"))) ВнеОбАРезИсслед.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1120", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Period2"))) ВнеОбАРезИсслед.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1120", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1120", "Period3"))) ВнеОбАРезИсслед.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1120", "Period3")).ToString();
                                ВнеОбА.РезИсслед = ВнеОбАРезИсслед;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАНеМатПоискАкт = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Poya"))) ВнеОбАНеМатПоискАкт.Пояснения = GetFields.gettabledata(Document, Processing, "1130", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Period1"))) ВнеОбАНеМатПоискАкт.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1130", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Period2"))) ВнеОбАНеМатПоискАкт.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1130", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1130", "Period3"))) ВнеОбАНеМатПоискАкт.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1130", "Period3")).ToString();
                                ВнеОбА.НеМатПоискАкт = ВнеОбАНеМатПоискАкт;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАМатПоискАкт = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Poya"))) ВнеОбАМатПоискАкт.Пояснения = GetFields.gettabledata(Document, Processing, "1140", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Period1"))) ВнеОбАМатПоискАкт.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1140", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Period2"))) ВнеОбАМатПоискАкт.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1140", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1140", "Period3"))) ВнеОбАМатПоискАкт.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1140", "Period3")).ToString();
                                ВнеОбА.МатПоискАкт = ВнеОбАМатПоискАкт;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАОснСр = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Poya"))) ВнеОбАОснСр.Пояснения = GetFields.gettabledata(Document, Processing, "1150", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Period1"))) ВнеОбАОснСр.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1150", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Period2"))) ВнеОбАОснСр.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1150", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1150", "Period3"))) ВнеОбАОснСр.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1150", "Period3")).ToString();
                                ВнеОбА.ОснСр = ВнеОбАОснСр;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАВлМатЦен = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Poya"))) ВнеОбАВлМатЦен.Пояснения = GetFields.gettabledata(Document, Processing, "1160", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Period1"))) ВнеОбАВлМатЦен.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1160", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Period2"))) ВнеОбАВлМатЦен.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1160", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1160", "Period3"))) ВнеОбАВлМатЦен.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1160", "Period3")).ToString();
                                ВнеОбА.ВлМатЦен = ВнеОбАВлМатЦен;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАФинВлож = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Poya"))) ВнеОбАФинВлож.Пояснения = GetFields.gettabledata(Document, Processing, "1170", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Period1"))) ВнеОбАФинВлож.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1170", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Period2"))) ВнеОбАФинВлож.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1170", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1170", "Period3"))) ВнеОбАФинВлож.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1170", "Period3")).ToString();
                                ВнеОбА.ФинВлож = ВнеОбАФинВлож;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАОтлНалАкт = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Poya"))) ВнеОбАОтлНалАкт.Пояснения = GetFields.gettabledata(Document, Processing, "1180", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Period1"))) ВнеОбАОтлНалАкт.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1180", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Period2"))) ВнеОбАОтлНалАкт.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1180", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1180", "Period3"))) ВнеОбАОтлНалАкт.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1180", "Period3")).ToString();
                                ВнеОбА.ОтлНалАкт = ВнеОбАОтлНалАкт;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Period3"))))
                            {
                                ОПП_ВПТип ВнеОбАПрочВнеОбА = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Poya"))) ВнеОбАПрочВнеОбА.Пояснения = GetFields.gettabledata(Document, Processing, "1190", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Period1"))) ВнеОбАПрочВнеОбА.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1190", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Period2"))) ВнеОбАПрочВнеОбА.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1190", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1190", "Period3"))) ВнеОбАПрочВнеОбА.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1190", "Period3")).ToString();
                                ВнеОбА.ПрочВнеОбА = ВнеОбАПрочВнеОбА;
                            }
                            Aктив.ВнеОбА = ВнеОбА;


                        }
                        #endregion ВнеОбА

                        #region ОбА
                        if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Period3"))) || GetFields.hastabledata(Document, Processing, "1210") || GetFields.hastabledata(Document, Processing, "1220") || GetFields.hastabledata(Document, Processing, "1230") || GetFields.hastabledata(Document, Processing, "1240") || GetFields.hastabledata(Document, Processing, "1250") || GetFields.hastabledata(Document, Processing, "1260"))
                        {
                            var ОбА = new ФайлДокументБалансАктивОбА();
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Poya"))) { ОбА.Пояснения = GetFields.gettabledata(Document, Processing, "1200", "Poya"); }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Period1"))) { ОбА.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1200", "Period1")).ToString(); } else { ОбА.СумОтч = "0"; }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Period2"))) { ОбА.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1200", "Period2")).ToString(); } else { ОбА.СумПрдщ = "0"; }
                            if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1200", "Period3"))) { ОбА.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1200", "Period3")).ToString(); } else { ОбА.СумПрдшв = "0"; }

                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Period3"))))
                            {
                                ОПП_ВПТип ОбАЗапасы = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Poya"))) ОбАЗапасы.Пояснения = GetFields.gettabledata(Document, Processing, "1210", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Period1"))) ОбАЗапасы.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1210", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Period2"))) ОбАЗапасы.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1210", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1210", "Period3"))) ОбАЗапасы.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1210", "Period3")).ToString();
                                ОбА.Запасы = ОбАЗапасы;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Period3"))))
                            {
                                ОПП_ВПТип ОбАНДСПриобрЦен = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Poya"))) ОбАНДСПриобрЦен.Пояснения = GetFields.gettabledata(Document, Processing, "1220", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Period1"))) ОбАНДСПриобрЦен.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1220", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Period2"))) ОбАНДСПриобрЦен.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1220", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1220", "Period3"))) ОбАНДСПриобрЦен.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1220", "Period3")).ToString();
                                ОбА.НДСПриобрЦен = ОбАНДСПриобрЦен;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Period3"))))
                            {
                                ОПП_ВПТип ОбАДебЗад = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Poya"))) ОбАДебЗад.Пояснения = GetFields.gettabledata(Document, Processing, "1230", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Period1"))) ОбАДебЗад.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1230", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Period2"))) ОбАДебЗад.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1230", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1230", "Period3"))) ОбАДебЗад.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1230", "Period3")).ToString();
                                ОбА.ДебЗад = ОбАДебЗад;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Period3"))))
                            {
                                ОПП_ВПТип ОбАФинВлож = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Poya"))) ОбАФинВлож.Пояснения = GetFields.gettabledata(Document, Processing, "1240", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Period1"))) ОбАФинВлож.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1240", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Period2"))) ОбАФинВлож.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1240", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1240", "Period3"))) ОбАФинВлож.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1240", "Period3")).ToString();
                                ОбА.ФинВлож = ОбАФинВлож;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Period3"))))
                            {
                                ОПП_ВПТип ОбАДенежнСр = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Poya"))) ОбАДенежнСр.Пояснения = GetFields.gettabledata(Document, Processing, "1250", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Period1"))) ОбАДенежнСр.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1250", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Period2"))) ОбАДенежнСр.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1250", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1250", "Period3"))) ОбАДенежнСр.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1250", "Period3")).ToString();
                                ОбА.ДенежнСр = ОбАДенежнСр;
                            }
                            if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Poya"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Period2"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Period3"))))
                            {
                                ОПП_ВПТип ОбАПрочОбА = new ОПП_ВПТип();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Poya"))) ОбАПрочОбА.Пояснения = GetFields.gettabledata(Document, Processing, "1260", "Poya");
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Period1"))) ОбАПрочОбА.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1260", "Period1")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Period2"))) ОбАПрочОбА.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1260", "Period2")).ToString();
                                if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "1260", "Period3"))) ОбАПрочОбА.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1260", "Period3")).ToString();
                                ОбА.ПрочОбА = ОбАПрочОбА;
                            }
                            Aктив.ОбА = ОбА;
                        }
                        Баланс.Актив = Aктив;
                    }
                    #endregion ОбА




                    #endregion Актив

                    #region Пассив
                    if ((GetFields.gettabledata(Document, Processing, "1700", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1700", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1700", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1700", "Period3") != ""))
                    {
                        var Пассив = new ФайлДокументБалансПассив();
                        if (GetFields.gettabledata(Document, Processing, "1700", "Poya") != "") Пассив.Пояснения = GetFields.gettabledata(Document, Processing, "1700", "Poya");
                        if (GetFields.gettabledata(Document, Processing, "1700", "Period1") != "") Пассив.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1700", "Period1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "1700", "Period2") != "") Пассив.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1700", "Period2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "1700", "Period3") != "") Пассив.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1700", "Period3")).ToString();

                        #region КапРез
                        if ((GetFields.gettabledata(Document, Processing, "1300", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1300", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1300", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1300", "Period3") != "") || GetFields.hastabledata(Document, Processing, "1310") || GetFields.hastabledata(Document, Processing, "1320") || GetFields.hastabledata(Document, Processing, "1340") || GetFields.hastabledata(Document, Processing, "1350") || GetFields.hastabledata(Document, Processing, "1360") || GetFields.hastabledata(Document, Processing, "1370"))
                        {
                            var КапРез = new ФайлДокументБалансПассивКапРез();
                            if (GetFields.gettabledata(Document, Processing, "1300", "Poya") != "") { КапРез.Пояснения = GetFields.gettabledata(Document, Processing, "1300", "Poya"); }
                            if (GetFields.gettabledata(Document, Processing, "1300", "Period1") != "") { КапРез.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1300", "Period1")).ToString(); } else { КапРез.СумОтч = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "1300", "Period2") != "") { КапРез.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1300", "Period2")).ToString(); } else { КапРез.СумПрдщ = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "1300", "Period3") != "") { КапРез.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1300", "Period3")).ToString(); } else { КапРез.СумПрдшв = "0"; }

                            if ((GetFields.gettabledata(Document, Processing, "1310", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1310", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1310", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1310", "Period3") != ""))
                            {
                                ОПП_ВПТип КапРезУставКапитал = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1310", "Poya") != "") КапРезУставКапитал.Пояснения = GetFields.gettabledata(Document, Processing, "1310", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1310", "Period1") != "") КапРезУставКапитал.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1310", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1310", "Period2") != "") КапРезУставКапитал.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1310", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1310", "Period3") != "") КапРезУставКапитал.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1310", "Period3")).ToString();
                                КапРез.УставКапитал = КапРезУставКапитал;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1320", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1320", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1320", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1320", "Period3") != ""))
                            {
                                ОПП_ВПТип КапРезСобствАкции = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1320", "Poya") != "") КапРезСобствАкции.Пояснения = GetFields.gettabledata(Document, Processing, "1320", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1320", "Period1") != "") КапРезСобствАкции.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1320", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1320", "Period2") != "") КапРезСобствАкции.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1320", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1320", "Period3") != "") КапРезСобствАкции.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1320", "Period3")).ToString();
                                КапРез.СобствАкции = КапРезСобствАкции;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1340", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1340", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1340", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1340", "Period3") != ""))
                            {
                                ОПП_ВПТип КапРезПереоцВнеОбА = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1340", "Poya") != "") КапРезПереоцВнеОбА.Пояснения = GetFields.gettabledata(Document, Processing, "1340", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1340", "Period1") != "") КапРезПереоцВнеОбА.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1340", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1340", "Period2") != "") КапРезПереоцВнеОбА.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1340", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1340", "Period3") != "") КапРезПереоцВнеОбА.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1340", "Period3")).ToString();
                                КапРез.ПереоцВнеОбА = КапРезПереоцВнеОбА;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1350", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1350", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1350", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1350", "Period3") != ""))
                            {
                                ОПП_ВПТип КапРезДобКапитал = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1350", "Poya") != "") КапРезДобКапитал.Пояснения = GetFields.gettabledata(Document, Processing, "1350", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1350", "Period1") != "") КапРезДобКапитал.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1350", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1350", "Period2") != "") КапРезДобКапитал.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1350", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1350", "Period3") != "") КапРезДобКапитал.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1350", "Period3")).ToString();
                                КапРез.ДобКапитал = КапРезДобКапитал;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1360", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1360", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1360", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1360", "Period3") != ""))
                            {
                                ОПП_ВПТип КапРезРезКапитал = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1360", "Poya") != "") КапРезРезКапитал.Пояснения = GetFields.gettabledata(Document, Processing, "1360", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1360", "Period1") != "") КапРезРезКапитал.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1360", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1360", "Period2") != "") КапРезРезКапитал.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1360", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1360", "Period3") != "") КапРезРезКапитал.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1360", "Period3")).ToString();
                                КапРез.РезКапитал = КапРезРезКапитал;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1370", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1370", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1370", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1370", "Period3") != ""))
                            {
                                ОПП_ВПТип КапРезНераспПриб = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1370", "Poya") != "") КапРезНераспПриб.Пояснения = GetFields.gettabledata(Document, Processing, "1370", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1370", "Period1") != "") КапРезНераспПриб.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1370", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1370", "Period2") != "") КапРезНераспПриб.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1370", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1370", "Period3") != "") КапРезНераспПриб.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1370", "Period3")).ToString();
                                КапРез.НераспПриб = КапРезНераспПриб;
                            }
                            Пассив.Item = КапРез;
                        }
                        #endregion

                        #region ДолгосрОбяз
                        if ((GetFields.gettabledata(Document, Processing, "1400", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1400", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1400", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1400", "Period3") != "") || GetFields.hastabledata(Document, Processing, "1410") || GetFields.hastabledata(Document, Processing, "1420") || GetFields.hastabledata(Document, Processing, "1430") || GetFields.hastabledata(Document, Processing, "1450"))
                        {
                            var ДолгосрОбяз = new ФайлДокументБалансПассивДолгосрОбяз();
                            if (GetFields.gettabledata(Document, Processing, "1400", "Poya") != "") { ДолгосрОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1400", "Poya"); }
                            if (GetFields.gettabledata(Document, Processing, "1400", "Period1") != "") { ДолгосрОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1400", "Period1")).ToString(); } else { ДолгосрОбяз.СумОтч = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "1400", "Period2") != "") { ДолгосрОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1400", "Period2")).ToString(); } else { ДолгосрОбяз.СумПрдщ = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "1400", "Period3") != "") { ДолгосрОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1400", "Period3")).ToString(); } else { ДолгосрОбяз.СумПрдшв = "0"; }

                            if ((GetFields.gettabledata(Document, Processing, "1410", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1410", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1410", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1410", "Period3") != ""))
                            {
                                ОПП_ВПТип ДолгосрОбязЗаемСредств = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1410", "Poya") != "") ДолгосрОбязЗаемСредств.Пояснения = GetFields.gettabledata(Document, Processing, "1410", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1410", "Period1") != "") ДолгосрОбязЗаемСредств.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1410", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1410", "Period2") != "") ДолгосрОбязЗаемСредств.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1410", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1410", "Period3") != "") ДолгосрОбязЗаемСредств.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1410", "Period3")).ToString();
                                ДолгосрОбяз.ЗаемСредств = ДолгосрОбязЗаемСредств;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1420", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1420", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1420", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1420", "Period3") != ""))
                            {
                                ОПП_ВПТип ДолгосрОбязОтложНалОбяз = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1420", "Poya") != "") ДолгосрОбязОтложНалОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1420", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1420", "Period1") != "") ДолгосрОбязОтложНалОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1420", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1420", "Period2") != "") ДолгосрОбязОтложНалОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1420", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1420", "Period3") != "") ДолгосрОбязОтложНалОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1420", "Period3")).ToString();
                                ДолгосрОбяз.ОтложНалОбяз = ДолгосрОбязОтложНалОбяз;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1430", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1430", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1430", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1430", "Period3") != ""))
                            {
                                ОПП_ВПТип ДолгосрОбязОценОбяз = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1430", "Poya") != "") ДолгосрОбязОценОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1430", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1430", "Period1") != "") ДолгосрОбязОценОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1430", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1430", "Period2") != "") ДолгосрОбязОценОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1430", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1430", "Period3") != "") ДолгосрОбязОценОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1430", "Period3")).ToString();
                                ДолгосрОбяз.ОценОбяз = ДолгосрОбязОценОбяз;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1450", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1450", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1450", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1450", "Period3") != ""))
                            {
                                ОПП_ВПТип ДолгосрОбязПрочОбяз = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1450", "Poya") != "") ДолгосрОбязПрочОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1450", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1450", "Period1") != "") ДолгосрОбязПрочОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1450", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1450", "Period2") != "") ДолгосрОбязПрочОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1450", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1450", "Period3") != "") ДолгосрОбязПрочОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1450", "Period3")).ToString();
                                ДолгосрОбяз.ПрочОбяз = ДолгосрОбязПрочОбяз;
                            }

                            Пассив.ДолгосрОбяз = ДолгосрОбяз;
                        }
                        #endregion

                        #region КраткосрОбяз
                        if ((GetFields.gettabledata(Document, Processing, "1500", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1500", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1500", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1500", "Period3") != "") || GetFields.hastabledata(Document, Processing, "1510") || GetFields.hastabledata(Document, Processing, "1520") || GetFields.hastabledata(Document, Processing, "1530") || GetFields.hastabledata(Document, Processing, "1540") || GetFields.hastabledata(Document, Processing, "1550"))
                        {
                            var КраткосрОбяз = new ФайлДокументБалансПассивКраткосрОбяз();
                            if (GetFields.gettabledata(Document, Processing, "1500", "Poya") != "") { КраткосрОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1500", "Poya"); }
                            if (GetFields.gettabledata(Document, Processing, "1500", "Period1") != "") { КраткосрОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1500", "Period1")).ToString(); } else { КраткосрОбяз.СумОтч = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "1500", "Period2") != "") { КраткосрОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1500", "Period2")).ToString(); } else { КраткосрОбяз.СумПрдщ = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "1500", "Period3") != "") { КраткосрОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1500", "Period3")).ToString(); } else { КраткосрОбяз.СумПрдшв = "0"; }

                            if ((GetFields.gettabledata(Document, Processing, "1510", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1510", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1510", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1510", "Period3") != ""))
                            {
                                ОПП_ВПТип КраткосрОбязЗаемСредств = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1510", "Poya") != "") КраткосрОбязЗаемСредств.Пояснения = GetFields.gettabledata(Document, Processing, "1510", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1510", "Period1") != "") КраткосрОбязЗаемСредств.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1510", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1510", "Period2") != "") КраткосрОбязЗаемСредств.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1510", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1510", "Period3") != "") КраткосрОбязЗаемСредств.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1510", "Period3")).ToString();
                                КраткосрОбяз.ЗаемСредств = КраткосрОбязЗаемСредств;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1520", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1520", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1520", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1520", "Period3") != ""))
                            {
                                ОПП_ВПТип КраткосрОбязКредитЗадолж = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1520", "Poya") != "") КраткосрОбязКредитЗадолж.Пояснения = GetFields.gettabledata(Document, Processing, "1520", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1520", "Period1") != "") КраткосрОбязКредитЗадолж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1520", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1520", "Period2") != "") КраткосрОбязКредитЗадолж.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1520", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1520", "Period3") != "") КраткосрОбязКредитЗадолж.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1520", "Period3")).ToString();
                                КраткосрОбяз.КредитЗадолж = КраткосрОбязКредитЗадолж;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1530", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1530", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1530", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1530", "Period3") != ""))
                            {
                                ОПП_ВПТип КраткосрОбязДоходБудущ = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1530", "Poya") != "") КраткосрОбязДоходБудущ.Пояснения = GetFields.gettabledata(Document, Processing, "1530", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1530", "Period1") != "") КраткосрОбязДоходБудущ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1530", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1530", "Period2") != "") КраткосрОбязДоходБудущ.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1530", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1530", "Period3") != "") КраткосрОбязДоходБудущ.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1530", "Period3")).ToString();
                                КраткосрОбяз.ДоходБудущ = КраткосрОбязДоходБудущ;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1540", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1540", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1540", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1540", "Period3") != ""))
                            {
                                ОПП_ВПТип КраткосрОбязОценОбяз = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1540", "Poya") != "") КраткосрОбязОценОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1540", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1540", "Period1") != "") КраткосрОбязОценОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1540", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1540", "Period2") != "") КраткосрОбязОценОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1540", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1540", "Period3") != "") КраткосрОбязОценОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1540", "Period3")).ToString();
                                КраткосрОбяз.ОценОбяз = КраткосрОбязОценОбяз;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "1550", "Poya") != "") || (GetFields.gettabledata(Document, Processing, "1550", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "1550", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "1550", "Period3") != ""))
                            {
                                ОПП_ВПТип КраткосрОбязПрочОбяз = new ОПП_ВПТип();
                                if (GetFields.gettabledata(Document, Processing, "1550", "Poya") != "") КраткосрОбязПрочОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "1550", "Poya");
                                if (GetFields.gettabledata(Document, Processing, "1550", "Period1") != "") КраткосрОбязПрочОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1550", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1550", "Period2") != "") КраткосрОбязПрочОбяз.СумПрдщ = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1550", "Period2")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "1550", "Period3") != "") КраткосрОбязПрочОбяз.СумПрдшв = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "1550", "Period3")).ToString();
                                КраткосрОбяз.ПрочОбяз = КраткосрОбязПрочОбяз;
                            }
                            Пассив.КраткосрОбяз = КраткосрОбяз;
                        }
                        #endregion
                        Баланс.Пассив = Пассив;
                    }
                    #endregion Пассив

                    Баланс.ОКУД = БалансОКУД;
                    Баланс.ОКЕИ = ФайлДокументБалансОКЕИ.Item384;
                    Баланс.Период = БалансПериод;
                    Document.Property("UncertainSymbolsCount");
                    Баланс.ПроцентРаспознавания = (100 - (Math.Round(Convert.ToDecimal(Convert.ToInt32(Document.SymbolsForVerificationCount.ToString()) * 100 / Convert.ToInt32(Document.TotalSymbolsCount.ToString()))))).ToString();

                    DateTime dateValue2;

                    if (GetFields.GetField(Document, Processing, "Date") != "" && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out dateValue2)) { Баланс.ОтчетГод = dateValue2.Year.ToString(); /*Convert.ToDateTime(GetFields.GetField(Document, Processing, "Date")).Year.ToString();*/} else { Баланс.ОтчетГод = DateTime.Now.Year.ToString(); }


                    ФайлДокументБаланс[] _Баланс = { Баланс };
                    Документ.Баланс = _Баланс;

                }
                #endregion Документ_Баланс

                #region Документ_ПрибУб
                if (Document.DefinitionName == "3_37_1_Финансовый отчет" || Document.DefinitionName == "3_37_3_Бухгалтерский баланс ОКУД 0710099" || Document.DefinitionName == "3_37_4_Упрощенная бухгалтерская (финансовая) отчетность")
                {

                    var ПрибУб = new ФайлДокументПрибУб();
                    var ПрибУбОКУД = new ФайлДокументПрибУбОКУД();
                    var ПрибУбОКЕИ = new ФайлДокументПрибУбОКЕИ();
                    var ПрибУбПериод = new ПериодКварталТип();
                    ПрибУбПериод = PeriodQuarterType(GetFields.GetField(Document, Processing, "Date"));

                    if ((GetFields.gettabledata(Document, Processing, "2110", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2110", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2110", "Period2") != ""))
                    {
                        var ПрибУбВыруч = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2110", "poya") != "") ПрибУбВыруч.Пояснения = GetFields.gettabledata(Document, Processing, "2110", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2110", "Period1") != "") ПрибУбВыруч.СумОтч = GetFields.gettabledata(Document, Processing, "2110", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2110", "Period2") != "") ПрибУбВыруч.СумПред = GetFields.gettabledata(Document, Processing, "2110", "Period2");
                        ПрибУб.Выруч = ПрибУбВыруч;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "2120", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2120", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2120", "Period2") != ""))
                    {
                        var ПрибУбСебестПрод = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2120", "poya") != "") ПрибУбСебестПрод.Пояснения = GetFields.gettabledata(Document, Processing, "2120", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2120", "Period1") != "") ПрибУбСебестПрод.СумОтч = GetFields.gettabledata(Document, Processing, "2120", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2120", "Period2") != "") ПрибУбСебестПрод.СумПред = GetFields.gettabledata(Document, Processing, "2120", "Period2");
                        ПрибУб.СебестПрод = ПрибУбСебестПрод;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2100", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2100", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2100", "Period2") != ""))
                    {
                        var ПрибУбВаловаяПрибыль = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2100", "poya") != "") ПрибУбВаловаяПрибыль.Пояснения = GetFields.gettabledata(Document, Processing, "2100", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2100", "Period1") != "") ПрибУбВаловаяПрибыль.СумОтч = GetFields.gettabledata(Document, Processing, "2100", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2100", "Period2") != "") ПрибУбВаловаяПрибыль.СумПред = GetFields.gettabledata(Document, Processing, "2100", "Period2");
                        ПрибУб.ВаловаяПрибыль = ПрибУбВаловаяПрибыль;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2210", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2210", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2210", "Period2") != ""))
                    {
                        var ПрибУбКомРасход = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2210", "poya") != "") ПрибУбКомРасход.Пояснения = GetFields.gettabledata(Document, Processing, "2210", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2210", "Period1") != "") ПрибУбКомРасход.СумОтч = GetFields.gettabledata(Document, Processing, "2210", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2210", "Period2") != "") ПрибУбКомРасход.СумПред = GetFields.gettabledata(Document, Processing, "2210", "Period2");
                        ПрибУб.КомРасход = ПрибУбКомРасход;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2220", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2220", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2220", "Period2") != ""))
                    {
                        var ПрибУбУпрРасход = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2220", "poya") != "") ПрибУбУпрРасход.Пояснения = GetFields.gettabledata(Document, Processing, "2220", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2220", "Period1") != "") ПрибУбУпрРасход.СумОтч = GetFields.gettabledata(Document, Processing, "2220", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2220", "Period2") != "") ПрибУбУпрРасход.СумПред = GetFields.gettabledata(Document, Processing, "2220", "Period2");
                        ПрибУб.УпрРасход = ПрибУбУпрРасход;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2200", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2200", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2200", "Period2") != ""))
                    {
                        var ПрибУбПрибПрод = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2200", "poya") != "") ПрибУбПрибПрод.Пояснения = GetFields.gettabledata(Document, Processing, "2200", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2200", "Period1") != "") ПрибУбПрибПрод.СумОтч = GetFields.gettabledata(Document, Processing, "2200", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2200", "Period2") != "") ПрибУбПрибПрод.СумПред = GetFields.gettabledata(Document, Processing, "2200", "Period2");
                        ПрибУб.ПрибПрод = ПрибУбПрибПрод;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2310", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2310", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2310", "Period2") != ""))
                    {
                        var ПрибУбДоходОтУчаст = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2310", "poya") != "") ПрибУбДоходОтУчаст.Пояснения = GetFields.gettabledata(Document, Processing, "2310", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2310", "Period1") != "") ПрибУбДоходОтУчаст.СумОтч = GetFields.gettabledata(Document, Processing, "2310", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2310", "Period2") != "") ПрибУбДоходОтУчаст.СумПред = GetFields.gettabledata(Document, Processing, "2310", "Period2");
                        ПрибУб.ДоходОтУчаст = ПрибУбДоходОтУчаст;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2320", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2320", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2320", "Period2") != ""))
                    {
                        var ПрибУбПроцПолуч = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2320", "poya") != "") ПрибУбПроцПолуч.Пояснения = GetFields.gettabledata(Document, Processing, "2320", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2320", "Period1") != "") ПрибУбПроцПолуч.СумОтч = GetFields.gettabledata(Document, Processing, "2320", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2320", "Period2") != "") ПрибУбПроцПолуч.СумПред = GetFields.gettabledata(Document, Processing, "2320", "Period2");
                        ПрибУб.ПроцПолуч = ПрибУбПроцПолуч;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2330", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2330", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2330", "Period2") != ""))
                    {
                        var ПрибУбПроцУпл = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2330", "poya") != "") ПрибУбПроцУпл.Пояснения = GetFields.gettabledata(Document, Processing, "2330", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2330", "Period1") != "") ПрибУбПроцУпл.СумОтч = GetFields.gettabledata(Document, Processing, "2330", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2330", "Period2") != "") ПрибУбПроцУпл.СумПред = GetFields.gettabledata(Document, Processing, "2330", "Period2");
                        ПрибУб.ПроцУпл = ПрибУбПроцУпл;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2340", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2340", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2340", "Period2") != ""))
                    {
                        var ПрибУбПрочДоход = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2340", "poya") != "") ПрибУбПрочДоход.Пояснения = GetFields.gettabledata(Document, Processing, "2340", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2340", "Period1") != "") ПрибУбПрочДоход.СумОтч = GetFields.gettabledata(Document, Processing, "2340", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2340", "Period2") != "") ПрибУбПрочДоход.СумПред = GetFields.gettabledata(Document, Processing, "2340", "Period2");
                        ПрибУб.ПрочДоход = ПрибУбПрочДоход;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2350", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2350", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2350", "Period2") != ""))
                    {
                        var ПрибУбПрочРасход = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2350", "poya") != "") ПрибУбПрочРасход.Пояснения = GetFields.gettabledata(Document, Processing, "2350", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2350", "Period1") != "") ПрибУбПрочРасход.СумОтч = GetFields.gettabledata(Document, Processing, "2350", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2350", "Period2") != "") ПрибУбПрочРасход.СумПред = GetFields.gettabledata(Document, Processing, "2350", "Period2");
                        ПрибУб.ПрочРасход = ПрибУбПрочРасход;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2300", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2300", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2300", "Period2") != ""))
                    {
                        var ПрибУбПрибУбДоНал = new ОтчПредОДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2300", "poya") != "") ПрибУбПрибУбДоНал.Пояснения = GetFields.gettabledata(Document, Processing, "2300", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2300", "Period1") != "") ПрибУбПрибУбДоНал.СумОтч = GetFields.gettabledata(Document, Processing, "2300", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2300", "Period2") != "") ПрибУбПрибУбДоНал.СумПред = GetFields.gettabledata(Document, Processing, "2300", "Period2");
                        ПрибУб.ПрибУбДоНал = ПрибУбПрибУбДоНал;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2410", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2410", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2410", "Period2") != ""))
                    {
                        var ПрибУбТекНалПриб = new ОтчПредНТип();
                        if (GetFields.gettabledata(Document, Processing, "2410", "poya") != "") ПрибУбТекНалПриб.Пояснения = GetFields.gettabledata(Document, Processing, "2410", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2410", "Period1") != "") ПрибУбТекНалПриб.СумОтч = GetFields.gettabledata(Document, Processing, "2410", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2410", "Period2") != "") ПрибУбТекНалПриб.СумПред = GetFields.gettabledata(Document, Processing, "2410", "Period2");
                        ПрибУб.ТекНалПриб = ПрибУбТекНалПриб;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2421", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2421", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2421", "Period2") != ""))
                    {
                        var ПрибУбПостНалОбяз = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2421", "poya") != "") ПрибУбПостНалОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "2421", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2421", "Period1") != "") ПрибУбПостНалОбяз.СумОтч = GetFields.gettabledata(Document, Processing, "2421", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2421", "Period2") != "") ПрибУбПостНалОбяз.СумПред = GetFields.gettabledata(Document, Processing, "2421", "Period2");
                        ПрибУб.ПостНалОбяз = ПрибУбПостНалОбяз;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2430", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2430", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2430", "Period2") != ""))
                    {
                        var ПрибУбИзмНалОбяз = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2430", "poya") != "") ПрибУбИзмНалОбяз.Пояснения = GetFields.gettabledata(Document, Processing, "2430", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2430", "Period1") != "") ПрибУбИзмНалОбяз.СумОтч = GetFields.gettabledata(Document, Processing, "2430", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2430", "Period2") != "") ПрибУбИзмНалОбяз.СумПред = GetFields.gettabledata(Document, Processing, "2430", "Period2");
                        ПрибУб.ИзмНалОбяз = ПрибУбИзмНалОбяз;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2450", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2450", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2450", "Period2") != ""))
                    {
                        var ПрибУбИзмНалАктив = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2450", "poya") != "") ПрибУбИзмНалАктив.Пояснения = GetFields.gettabledata(Document, Processing, "2450", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2450", "Period1") != "") ПрибУбИзмНалАктив.СумОтч = GetFields.gettabledata(Document, Processing, "2450", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2450", "Period2") != "") ПрибУбИзмНалАктив.СумПред = GetFields.gettabledata(Document, Processing, "2450", "Period2");
                        ПрибУб.ИзмНалАктив = ПрибУбИзмНалАктив;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2460", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2460", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2460", "Period2") != ""))
                    {
                        var ПрибУбПрочее = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2460", "poya") != "") ПрибУбПрочее.Пояснения = GetFields.gettabledata(Document, Processing, "2460", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2460", "Period1") != "") ПрибУбПрочее.СумОтч = GetFields.gettabledata(Document, Processing, "2460", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2460", "Period2") != "") ПрибУбПрочее.СумПред = GetFields.gettabledata(Document, Processing, "2460", "Period2");
                        ПрибУб.Прочее = ПрибУбПрочее;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2400", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2400", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2400", "Period2") != ""))
                    {
                        var ПрибУбЧистПрибУб = new ОтчПредОТип();
                        if (GetFields.gettabledata(Document, Processing, "2400", "poya") != "") ПрибУбЧистПрибУб.Пояснения = GetFields.gettabledata(Document, Processing, "2400", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2400", "Period1") != "") ПрибУбЧистПрибУб.СумОтч = GetFields.gettabledata(Document, Processing, "2400", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2400", "Period2") != "") ПрибУбЧистПрибУб.СумПред = GetFields.gettabledata(Document, Processing, "2400", "Period2");
                        ПрибУб.ЧистПрибУб = ПрибУбЧистПрибУб;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2510", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2510", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2510", "Period2") != ""))
                    {
                        var ПрибУбРезПрцВОАНеЧист = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2510", "poya") != "") ПрибУбРезПрцВОАНеЧист.Пояснения = GetFields.gettabledata(Document, Processing, "2510", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2510", "Period1") != "") ПрибУбРезПрцВОАНеЧист.СумОтч = GetFields.gettabledata(Document, Processing, "2510", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2510", "Period2") != "") ПрибУбРезПрцВОАНеЧист.СумПред = GetFields.gettabledata(Document, Processing, "2510", "Period2");
                        ПрибУб.РезПрцВОАНеЧист = ПрибУбРезПрцВОАНеЧист;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2520", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2520", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2520", "Period2") != ""))
                    {
                        var ПрибУбРезПрОпНеЧист = new ОтчПредНДопТип();
                        if (GetFields.gettabledata(Document, Processing, "2520", "poya") != "") ПрибУбРезПрОпНеЧист.Пояснения = GetFields.gettabledata(Document, Processing, "2520", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2520", "Period1") != "") ПрибУбРезПрОпНеЧист.СумОтч = GetFields.gettabledata(Document, Processing, "2520", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2520", "Period2") != "") ПрибУбРезПрОпНеЧист.СумПред = GetFields.gettabledata(Document, Processing, "2520", "Period2");
                        ПрибУб.РезПрОпНеЧист = ПрибУбРезПрОпНеЧист;
                    }

                    if ((GetFields.gettabledata(Document, Processing, "2500", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2500", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2500", "Period2") != ""))
                    {
                        var ПрибУбСовФинРез = new ОтчПредНТип();
                        if (GetFields.gettabledata(Document, Processing, "2500", "poya") != "") ПрибУбСовФинРез.Пояснения = GetFields.gettabledata(Document, Processing, "2500", "poya");
                        if (GetFields.gettabledata(Document, Processing, "2500", "Period1") != "") ПрибУбСовФинРез.СумОтч = GetFields.gettabledata(Document, Processing, "2500", "Period1");
                        if (GetFields.gettabledata(Document, Processing, "2500", "Period2") != "") ПрибУбСовФинРез.СумПред = GetFields.gettabledata(Document, Processing, "2500", "Period2");
                        ПрибУб.СовФинРез = ПрибУбСовФинРез;
                    }

                    #region Справочно
                    if ((GetFields.gettabledata(Document, Processing, "2900", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2900", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2900", "Period2") != "") || (GetFields.gettabledata(Document, Processing, "2910", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2910", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2910", "Period2") != ""))
                    {
                        var ПрибУбСправочно = new ФайлДокументПрибУбСправочно();

                        if ((GetFields.gettabledata(Document, Processing, "2900", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2900", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2900", "Period2") != ""))
                        {
                            var ПрибУбБазПрибылАкц = new ОтчПредНТип();
                            if (GetFields.gettabledata(Document, Processing, "2900", "poya") != "") ПрибУбБазПрибылАкц.Пояснения = GetFields.gettabledata(Document, Processing, "2900", "poya");
                            if (GetFields.gettabledata(Document, Processing, "2900", "Period1") != "") ПрибУбБазПрибылАкц.СумОтч = GetFields.gettabledata(Document, Processing, "2900", "Period1");
                            if (GetFields.gettabledata(Document, Processing, "2900", "Period2") != "") ПрибУбБазПрибылАкц.СумПред = GetFields.gettabledata(Document, Processing, "2900", "Period2");
                            ПрибУбСправочно.БазПрибылАкц = ПрибУбБазПрибылАкц;
                        }

                        if ((GetFields.gettabledata(Document, Processing, "2910", "poya") != "") || (GetFields.gettabledata(Document, Processing, "2910", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "2910", "Period2") != ""))
                        {
                            var ПрибУбРазводПрибылАкц = new ОтчПредНТип();
                            if (GetFields.gettabledata(Document, Processing, "2910", "poya") != "") ПрибУбРазводПрибылАкц.Пояснения = GetFields.gettabledata(Document, Processing, "2910", "poya");
                            if (GetFields.gettabledata(Document, Processing, "2910", "Period1") != "") ПрибУбРазводПрибылАкц.СумОтч = GetFields.gettabledata(Document, Processing, "2910", "Period1");
                            if (GetFields.gettabledata(Document, Processing, "2910", "Period2") != "") ПрибУбРазводПрибылАкц.СумПред = GetFields.gettabledata(Document, Processing, "2910", "Period2");
                            ПрибУбСправочно.РазводПрибылАкц = ПрибУбРазводПрибылАкц;
                        }
                        ПрибУб.Справочно = ПрибУбСправочно;
                    }
                    #endregion Справочно
                    ПрибУб.ОКЕИ = ФайлДокументПрибУбОКЕИ.Item384;
                    ПрибУб.ОКУД = ПрибУбОКУД;
                    ПрибУб.Период = ПрибУбПериод;
                    DateTime dateValue;
                    ПрибУб.ПроцентРаспознавания = (100 - (Math.Round(Convert.ToDecimal(Convert.ToInt32(Document.SymbolsForVerificationCount.ToString()) * 100 / Convert.ToInt32(Document.TotalSymbolsCount.ToString()))))).ToString();
                    if (GetFields.GetField(Document, Processing, "Date") != "" && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out dateValue)) { ПрибУб.ОтчетГод = dateValue.Year.ToString();/* Convert.ToDateTime(GetFields.GetField(Document, Processing, "Date")).Year.ToString(); */} else { ПрибУб.ОтчетГод = DateTime.Now.Year.ToString(); }
                    ФайлДокументПрибУб[] _ПрибУб = { ПрибУб };
                    Документ.ПрибУб = _ПрибУб;
                }
                #endregion Документ_ПрибУб

                #region Документ_ДвижениеДен
                if (Document.DefinitionName == "3_105_Отчет о движении денежных средств")
                {
                    var ДвижениеДен = new ФайлДокументДвижениеДен();
                    var ДвижениеДенОКЕИ = new ФайлДокументДвижениеДенОКЕИ();
                    var ДвижениеДенОКУД = new ФайлДокументДвижениеДенОКУД();
                    var ДвижениеДенПериод = new ПериодГодТип();

                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Date")) && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out DateTime dateValue1)) { ДвижениеДен.ОтчетГод = dateValue1.Year.ToString(); /*Convert.ToDateTime(GetFields.GetField(Document, Processing, "Date")).Year.ToString();*/} else { ДвижениеДен.ОтчетГод = DateTime.Now.Year.ToString(); }


                    #region ДвижениеДенТекОпер
                    if ((GetFields.gettabledata(Document, Processing, "4100", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4100", "Period2") != ""))
                    {
                        var ДвижениеДенТекОпер = new ФайлДокументДвижениеДенТекОпер();

                        var ДвижениеДенТекОперСальдоТек = new ОПТип();
                        if (GetFields.gettabledata(Document, Processing, "4100", "Period1") != "") { ДвижениеДенТекОперСальдоТек.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4100", "Period1")).ToString(); } else { ДвижениеДенТекОперСальдоТек.СумОтч = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "4100", "Period2") != "") { ДвижениеДенТекОперСальдоТек.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4100", "Period2")).ToString(); } else { ДвижениеДенТекОперСальдоТек.СумПред = "0"; }

                        #region ДвижениеДенТекОперПоступ
                        if ((GetFields.gettabledata(Document, Processing, "4110", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4110", "Period2") != ""))
                        {
                            var ДвижениеДенТекОперПоступ = new ФайлДокументДвижениеДенТекОперПоступ(); //описал вне региона
                            if (GetFields.gettabledata(Document, Processing, "4110", "Period1") != "") ДвижениеДенТекОперПоступ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4110", "Period1")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "4110", "Period2") != "") ДвижениеДенТекОперПоступ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4110", "Period2")).ToString();

                            if ((GetFields.gettabledata(Document, Processing, "4111", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4111", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПоступПродПТРУ = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4111", "Period1") != "") ДвижениеДенТекОперПоступПродПТРУ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4111", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4111", "Period2") != "") ДвижениеДенТекОперПоступПродПТРУ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4111", "Period2")).ToString();
                                ДвижениеДенТекОперПоступ.ПродПТРУ = ДвижениеДенТекОперПоступПродПТРУ;
                            }

                            if ((GetFields.gettabledata(Document, Processing, "4112", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4112", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПоступАрЛицИнПлат = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4112", "Period1") != "") ДвижениеДенТекОперПоступАрЛицИнПлат.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4112", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4112", "Period2") != "") ДвижениеДенТекОперПоступАрЛицИнПлат.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4112", "Period2")).ToString();
                                ДвижениеДенТекОперПоступ.АрЛицИнПлат = ДвижениеДенТекОперПоступАрЛицИнПлат;
                            }

                            if ((GetFields.gettabledata(Document, Processing, "4113", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4113", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПоступПродФинВлож = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4113", "Period1") != "") ДвижениеДенТекОперПоступПродФинВлож.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4113", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4113", "Period2") != "") ДвижениеДенТекОперПоступПродФинВлож.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4113", "Period2")).ToString();
                                ДвижениеДенТекОперПоступ.ПродФинВлож = ДвижениеДенТекОперПоступПродФинВлож;
                            }

                            //var ДвижениеДенТекОперПоступВПокТекПост = new ВПокОПТип(); //массив
                            //ДвижениеДенТекОперПоступВПокТекПост.НаимПок = "";
                            //ДвижениеДенТекОперПоступВПокТекПост.СумОтч = "";
                            //ДвижениеДенТекОперПоступВПокТекПост.СумПред = "";

                            if ((GetFields.gettabledata(Document, Processing, "4119", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4119", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПоступПрочПоступ = new ОП_ДТип();
                                if (GetFields.gettabledata(Document, Processing, "4119", "Period1") != "") ДвижениеДенТекОперПоступПрочПоступ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4119", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4119", "Period2") != "") ДвижениеДенТекОперПоступПрочПоступ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4119", "Period2")).ToString();
                                ДвижениеДенТекОперПоступ.ПрочПоступ = ДвижениеДенТекОперПоступПрочПоступ;
                            }

                            //var ДвижениеДенТекОперПоступПрочПоступВПокОП = new ВПокОПТип(); //массив
                            //ДвижениеДенТекОперПоступПрочПоступВПокОП.НаимПок = "";
                            //ДвижениеДенТекОперПоступПрочПоступВПокОП.СумОтч = "";
                            //ДвижениеДенТекОперПоступПрочПоступВПокОП.СумПред = "";

                            //ВПокОПТип[] _ДвижениеДенТекОперПоступПрочПоступВПокОП = { ДвижениеДенТекОперПоступПрочПоступВПокОП };
                            //ДвижениеДенТекОперПоступПрочПоступ.ВПокОП = _ДвижениеДенТекОперПоступПрочПоступВПокОП;

                            //ВПокОПТип[] _ДвижениеДенТекОперПоступВПокТекПост = { ДвижениеДенТекОперПоступВПокТекПост };
                            //ДвижениеДенТекОперПоступ.ВПокТекПост = _ДвижениеДенТекОперПоступВПокТекПост;



                            ДвижениеДенТекОпер.Поступ = ДвижениеДенТекОперПоступ;
                        }
                        #endregion Конец_ДвижениеДенТекОперПоступ

                        #region ДвижениеДенТекОперПлатеж
                        if ((GetFields.gettabledata(Document, Processing, "4120", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4120", "Period2") != ""))
                        {
                            var ДвижениеДенТекОперПлатеж = new ФайлДокументДвижениеДенТекОперПлатеж(); //описал вне региона
                            if (GetFields.gettabledata(Document, Processing, "4120", "Period1") != "") { ДвижениеДенТекОперПлатеж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4120", "Period1")).ToString(); } else { ДвижениеДенТекОперПлатеж.СумОтч = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "4120", "Period2") != "") { ДвижениеДенТекОперПлатеж.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4120", "Period2")).ToString(); } else { ДвижениеДенТекОперПлатеж.СумПред = "0"; }

                            if ((GetFields.gettabledata(Document, Processing, "4121", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4121", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПлатежПоставСМРУ = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4121", "Period1") != "") { ДвижениеДенТекОперПлатежПоставСМРУ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4121", "Period1")).ToString(); } else { ДвижениеДенТекОперПлатежПоставСМРУ.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4121", "Period2") != "") { ДвижениеДенТекОперПлатежПоставСМРУ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4121", "Period2")).ToString(); } else { ДвижениеДенТекОперПлатежПоставСМРУ.СумПред = "0"; }
                                ДвижениеДенТекОперПлатеж.ПоставСМРУ = ДвижениеДенТекОперПлатежПоставСМРУ;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4122", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4122", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПлатежОплатТрудРаб = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4122", "Period1") != "") { ДвижениеДенТекОперПлатежОплатТрудРаб.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4122", "Period1")).ToString(); } else { ДвижениеДенТекОперПлатежОплатТрудРаб.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4122", "Period2") != "") { ДвижениеДенТекОперПлатежОплатТрудРаб.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4122", "Period2")).ToString(); } else { ДвижениеДенТекОперПлатежОплатТрудРаб.СумПред = "0"; }
                                ДвижениеДенТекОперПлатеж.ОплатТрудРаб = ДвижениеДенТекОперПлатежОплатТрудРаб;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4123", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4123", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПлатежПроцДолгОбяз = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4123", "Period1") != "") { ДвижениеДенТекОперПлатежПроцДолгОбяз.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4123", "Period1")).ToString(); } else { ДвижениеДенТекОперПлатежПроцДолгОбяз.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4123", "Period2") != "") { ДвижениеДенТекОперПлатежПроцДолгОбяз.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4123", "Period2")).ToString(); } else { ДвижениеДенТекОперПлатежПроцДолгОбяз.СумПред = "0"; }
                                ДвижениеДенТекОперПлатеж.ПроцДолгОбяз = ДвижениеДенТекОперПлатежПроцДолгОбяз;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4124", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4124", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПлатежНалогПриб = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4124", "Period1") != "") { ДвижениеДенТекОперПлатежНалогПриб.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4124", "Period1")).ToString(); } else { ДвижениеДенТекОперПлатежНалогПриб.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4124", "Period2") != "") { ДвижениеДенТекОперПлатежНалогПриб.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4124", "Period2")).ToString(); } else { ДвижениеДенТекОперПлатежНалогПриб.СумПред = "0"; }
                                ДвижениеДенТекОперПлатеж.НалогПриб = ДвижениеДенТекОперПлатежНалогПриб;
                            }

                            //var ДвижениеДенТекОперПлатежВПокТекПлат = new ВПокОПТип(); //массив
                            //ДвижениеДенТекОперПлатежВПокТекПлат.НаимПок = "";
                            //ДвижениеДенТекОперПлатежВПокТекПлат.СумОтч = "";
                            //ДвижениеДенТекОперПлатежВПокТекПлат.СумПред = "";
                            if ((GetFields.gettabledata(Document, Processing, "4129", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4129", "Period2") != ""))
                            {
                                var ДвижениеДенТекОперПлатежПрочПлатеж = new ОП_ДТип();
                                if (GetFields.gettabledata(Document, Processing, "4129", "Period1") != "") { ДвижениеДенТекОперПлатежПрочПлатеж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4129", "Period1")).ToString(); } else { ДвижениеДенТекОперПлатежПрочПлатеж.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4129", "Period2") != "") { ДвижениеДенТекОперПлатежПрочПлатеж.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4129", "Period2")).ToString(); } else { ДвижениеДенТекОперПлатежПрочПлатеж.СумПред = "0"; }
                                ДвижениеДенТекОперПлатеж.ПрочПлатеж = ДвижениеДенТекОперПлатежПрочПлатеж;
                            }
                            //var ДвижениеДенТекОперПлатежПрочПлатежВПокОП = new ВПокОПТип(); //массив
                            //ДвижениеДенТекОперПлатежПрочПлатежВПокОП.НаимПок = "";
                            //ДвижениеДенТекОперПлатежПрочПлатежВПокОП.СумОтч = "";
                            //ДвижениеДенТекОперПлатежПрочПлатежВПокОП.СумПред = "";

                            //ВПокОПТип[] _ДвижениеДенТекОперПлатежПрочПлатежВПокОП = { ДвижениеДенТекОперПлатежПрочПлатежВПокОП };
                            //ДвижениеДенТекОперПлатежПрочПлатеж.ВПокОП = _ДвижениеДенТекОперПлатежПрочПлатежВПокОП;

                            ////ВПокОПТип[] _ДвижениеДенТекОперПлатежВПокТекПлат = { ДвижениеДенТекОперПлатежВПокТекПлат };
                            //ДвижениеДенТекОперПлатеж.ВПокТекПлат = _ДвижениеДенТекОперПлатежВПокТекПлат;

                            ДвижениеДенТекОпер.Платеж = ДвижениеДенТекОперПлатеж;
                        }
                        #endregion конец_ДвижениеДенТекОперПлатеж

                        ДвижениеДенТекОпер.СальдоТек = ДвижениеДенТекОперСальдоТек;
                        ДвижениеДен.ТекОпер = ДвижениеДенТекОпер;
                    }
                    #endregion Конец_ДвижениеДенТекОпер

                    #region ДвижениеДенИнвОпер
                    if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "4200", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "4200", "Period2"))))
                    {
                        var ДвижениеДенИнвОпер = new ФайлДокументДвижениеДенИнвОпер();
                        var ДвижениеДенИнвОперСальдоИнв = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "4200", "Period1"))){ДвижениеДенИнвОперСальдоИнв.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4200", "Period1")).ToString();}else{ ДвижениеДенИнвОперСальдоИнв.СумОтч = "0";}
                        if (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "4200", "Period2"))) {ДвижениеДенИнвОперСальдоИнв.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4200", "Period2")).ToString();} else { ДвижениеДенИнвОперСальдоИнв.СумПред = "0"; }

                        #region ДвижениеДенИнвОперПоступ
                        if ((!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "4210", "Period1"))) || (!string.IsNullOrEmpty(GetFields.gettabledata(Document, Processing, "4210", "Period2"))))
                        {
                            var ДвижениеДенИнвОперПоступ = new ФайлДокументДвижениеДенИнвОперПоступ(); //описал вне региона
                            if (GetFields.gettabledata(Document, Processing, "4210", "Period1") != "") {ДвижениеДенИнвОперПоступ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4210", "Period1")).ToString();} else { ДвижениеДенИнвОперПоступ.СумОтч = "0"; }
                            if (GetFields.gettabledata(Document, Processing, "4210", "Period2") != "") {ДвижениеДенИнвОперПоступ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4210", "Period2")).ToString();}

                            if ((GetFields.gettabledata(Document, Processing, "4211", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4211", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПоступПродВнАктив = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4211", "Period1") != "") {ДвижениеДенИнвОперПоступПродВнАктив.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4211", "Period1")).ToString();} else { ДвижениеДенИнвОперПоступПродВнАктив.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4211", "Period2") != "") {ДвижениеДенИнвОперПоступПродВнАктив.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4211", "Period2")).ToString();}
                                ДвижениеДенИнвОперПоступ.ПродВнАктив = ДвижениеДенИнвОперПоступПродВнАктив;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4212", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4212", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПоступПродАкцДр = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4212", "Period1") != "") {ДвижениеДенИнвОперПоступПродАкцДр.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4212", "Period1")).ToString();} else { ДвижениеДенИнвОперПоступПродАкцДр.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4212", "Period2") != "") {ДвижениеДенИнвОперПоступПродАкцДр.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4212", "Period2")).ToString();}
                                ДвижениеДенИнвОперПоступ.ПродАкцДр = ДвижениеДенИнвОперПоступПродАкцДр;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4213", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4213", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПоступВозврЗаймЦБ = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4213", "Period1") != "") {ДвижениеДенИнвОперПоступВозврЗаймЦБ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4213", "Period1")).ToString();} else { ДвижениеДенИнвОперПоступВозврЗаймЦБ.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4213", "Period2") != "") {ДвижениеДенИнвОперПоступВозврЗаймЦБ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4213", "Period2")).ToString();}
                                ДвижениеДенИнвОперПоступ.ВозврЗаймЦБ = ДвижениеДенИнвОперПоступВозврЗаймЦБ;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4214", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4214", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПоступДивПроц = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4214", "Period1") != "") {ДвижениеДенИнвОперПоступДивПроц.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4214", "Period1")).ToString();} else { ДвижениеДенИнвОперПоступДивПроц.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4214", "Period2") != "") {ДвижениеДенИнвОперПоступДивПроц.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4214", "Period2")).ToString();}
                                ДвижениеДенИнвОперПоступ.ДивПроц = ДвижениеДенИнвОперПоступДивПроц;
                            }

                            //var ДвижениеДенИнвОперПоступВПокИнвПост = new ВПокОПТип(); //Массик
                            //ДвижениеДенИнвОперПоступВПокИнвПост.НаимПок = "";
                            //ДвижениеДенИнвОперПоступВПокИнвПост.СумОтч = "";
                            //ДвижениеДенИнвОперПоступВПокИнвПост.СумПред = "";
                            if ((GetFields.gettabledata(Document, Processing, "4219", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4219", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПоступПрочПоступ = new ОП_ДТип();
                                if (GetFields.gettabledata(Document, Processing, "4219", "Period1") != "") {ДвижениеДенИнвОперПоступПрочПоступ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4219", "Period1")).ToString();} else { ДвижениеДенИнвОперПоступПрочПоступ.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4219", "Period2") != "") {ДвижениеДенИнвОперПоступПрочПоступ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4219", "Period2")).ToString();}
                                ДвижениеДенИнвОперПоступ.ПрочПоступ = ДвижениеДенИнвОперПоступПрочПоступ;
                            }
                            //var ДвижениеДенИнвОперПоступПрочПоступВПокОП = new ВПокОПТип(); //массив
                            //ДвижениеДенИнвОперПоступПрочПоступВПокОП.НаимПок = "";
                            //ДвижениеДенИнвОперПоступПрочПоступВПокОП.СумОтч = "";
                            //ДвижениеДенИнвОперПоступПрочПоступВПокОП.СумПред = "";

                            //Приравнивание
                            //ВПокОПТип[] _ДвижениеДенИнвОперПоступПрочПоступВПокОП = { ДвижениеДенИнвОперПоступПрочПоступВПокОП };
                            //ДвижениеДенИнвОперПоступПрочПоступ.ВПокОП = _ДвижениеДенИнвОперПоступПрочПоступВПокОП;

                            //ВПокОПТип[] _ДвижениеДенИнвОперПоступВПокИнвПост = { ДвижениеДенИнвОперПоступВПокИнвПост };
                            //ДвижениеДенИнвОперПоступ.ВПокИнвПост = _ДвижениеДенИнвОперПоступВПокИнвПост;





                            ДвижениеДенИнвОпер.Поступ = ДвижениеДенИнвОперПоступ;
                        }
                        #endregion //Конец ДвижениеДенИнвОперПоступ
                        #region ДвижениеДенИнвОперПлатеж
                        if ((GetFields.gettabledata(Document, Processing, "4220", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4220", "Period2") != ""))
                        {

                            var ДвижениеДенИнвОперПлатеж = new ФайлДокументДвижениеДенИнвОперПлатеж(); //описал вне региона
                            if (GetFields.gettabledata(Document, Processing, "4220", "Period1") != "") ДвижениеДенИнвОперПлатеж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4220", "Period1")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "4220", "Period2") != "") ДвижениеДенИнвОперПлатеж.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4220", "Period2")).ToString();

                            if ((GetFields.gettabledata(Document, Processing, "4221", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4221", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПлатежПриобрВнАктив = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4221", "Period1") != "") ДвижениеДенИнвОперПлатежПриобрВнАктив.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4221", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4221", "Period2") != "") ДвижениеДенИнвОперПлатежПриобрВнАктив.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4221", "Period2")).ToString();
                                ДвижениеДенИнвОперПлатеж.ПриобрВнАктив = ДвижениеДенИнвОперПлатежПриобрВнАктив;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4222", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4222", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПлатежПриобрАкцДр = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4222", "Period1") != "") ДвижениеДенИнвОперПлатежПриобрАкцДр.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4222", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4222", "Period2") != "") ДвижениеДенИнвОперПлатежПриобрАкцДр.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4222", "Period2")).ToString();
                                ДвижениеДенИнвОперПлатеж.ПриобрАкцДр = ДвижениеДенИнвОперПлатежПриобрАкцДр;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4223", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4223", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПлатежПриобрДолгЦБ = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4223", "Period1") != "") ДвижениеДенИнвОперПлатежПриобрДолгЦБ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4223", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4223", "Period2") != "") ДвижениеДенИнвОперПлатежПриобрДолгЦБ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4223", "Period2")).ToString();
                                ДвижениеДенИнвОперПлатеж.ПриобрДолгЦБ = ДвижениеДенИнвОперПлатежПриобрДолгЦБ;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4224", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4224", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПлатежПроцДолгОб = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4224", "Period1") != "") ДвижениеДенИнвОперПлатежПроцДолгОб.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4224", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4224", "Period2") != "") ДвижениеДенИнвОперПлатежПроцДолгОб.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4224", "Period2")).ToString();
                                ДвижениеДенИнвОперПлатеж.ПроцДолгОб = ДвижениеДенИнвОперПлатежПроцДолгОб;
                            }

                            //var ДвижениеДенИнвОперПлатежВПокИнвПлат = new ВПокОПТип(); //массив
                            //ДвижениеДенИнвОперПлатежВПокИнвПлат.НаимПок = "";
                            //ДвижениеДенИнвОперПлатежВПокИнвПлат.СумОтч = "";
                            //ДвижениеДенИнвОперПлатежВПокИнвПлат.СумПред = "";
                            if ((GetFields.gettabledata(Document, Processing, "4229", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4229", "Period2") != ""))
                            {
                                var ДвижениеДенИнвОперПлатежПрочПлатеж = new ОП_ДТип();
                                if (GetFields.gettabledata(Document, Processing, "4229", "Period1") != "") { ДвижениеДенИнвОперПлатежПрочПлатеж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4229", "Period1")).ToString(); } else { ДвижениеДенИнвОперПлатежПрочПлатеж.СумОтч = "0"; }
                                if (GetFields.gettabledata(Document, Processing, "4229", "Period2") != "") { ДвижениеДенИнвОперПлатежПрочПлатеж.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4229", "Period2")).ToString(); } else { ДвижениеДенИнвОперПлатежПрочПлатеж.СумПред = "0"; }
                                ДвижениеДенИнвОперПлатеж.ПрочПлатеж = ДвижениеДенИнвОперПлатежПрочПлатеж;
                            }

                            //var ДвижениеДенИнвОперПлатежПрочПлатежВПокОП = new ВПокОПТип(); //массив
                            //ДвижениеДенИнвОперПлатежПрочПлатежВПокОП.НаимПок = "";
                            //ДвижениеДенИнвОперПлатежПрочПлатежВПокОП.СумОтч = "";
                            //ДвижениеДенИнвОперПлатежПрочПлатежВПокОП.СумПред = "";

                            //Приравнивание
                            //ВПокОПТип[] _ДвижениеДенИнвОперПлатежПрочПлатежВПокОП = { ДвижениеДенИнвОперПлатежПрочПлатежВПокОП };
                            //ДвижениеДенИнвОперПлатежПрочПлатеж.ВПокОП = _ДвижениеДенИнвОперПлатежПрочПлатежВПокОП;

                            //ВПокОПТип[] _ДвижениеДенИнвОперПлатежВПокИнвПлат = { ДвижениеДенИнвОперПлатежВПокИнвПлат };
                            //ДвижениеДенИнвОперПлатеж.ВПокИнвПлат = _ДвижениеДенИнвОперПлатежВПокИнвПлат;



                            ДвижениеДенИнвОпер.Платеж = ДвижениеДенИнвОперПлатеж;
                        }
                        #endregion ДвижениеДенИнвОперПлатеж

                        ДвижениеДенИнвОпер.СальдоИнв = ДвижениеДенИнвОперСальдоИнв;
                        ДвижениеДен.ИнвОпер = ДвижениеДенИнвОпер;
                    }
                    #endregion конец_второй группы 

                    #region ДвижениеДенФинОпер
                    if ((GetFields.gettabledata(Document, Processing, "4300", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4300", "Period2") != ""))
                    {
                        var ДвижениеДенФинОпер = new ФайлДокументДвижениеДенФинОпер();
                        var ДвижениеДенФинОперСальдоФин = new ОПТип();
                        if (GetFields.gettabledata(Document, Processing, "4300", "Period1") != "") ДвижениеДенФинОперСальдоФин.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4300", "Period1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "4300", "Period2") != "") ДвижениеДенФинОперСальдоФин.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4300", "Period2")).ToString();

                        #region ДвижениеДенФинОперПоступ
                        if ((GetFields.gettabledata(Document, Processing, "4310", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4310", "Period2") != ""))
                        {
                            var ДвижениеДенФинОперПоступ = new ФайлДокументДвижениеДенФинОперПоступ(); //описал вне региона
                            if (GetFields.gettabledata(Document, Processing, "4310", "Period1") != "") ДвижениеДенФинОперПоступ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4310", "Period1")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "4310", "Period2") != "") ДвижениеДенФинОперПоступ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4310", "Period2")).ToString();

                            if ((GetFields.gettabledata(Document, Processing, "4311", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4311", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПоступКредЗайм = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4311", "Period1") != "") ДвижениеДенФинОперПоступКредЗайм.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4311", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4311", "Period2") != "") ДвижениеДенФинОперПоступКредЗайм.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4311", "Period2")).ToString();
                                ДвижениеДенФинОперПоступ.КредЗайм = ДвижениеДенФинОперПоступКредЗайм;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4312", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4312", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПоступВкладСоб = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4312", "Period1") != "") ДвижениеДенФинОперПоступВкладСоб.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4312", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4312", "Period2") != "") ДвижениеДенФинОперПоступВкладСоб.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4312", "Period2")).ToString();
                                ДвижениеДенФинОперПоступ.ВкладСоб = ДвижениеДенФинОперПоступВкладСоб;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4313", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4313", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПоступАкцДол = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4313", "Period1") != "") ДвижениеДенФинОперПоступАкцДол.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4313", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4313", "Period2") != "") ДвижениеДенФинОперПоступАкцДол.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4313", "Period2")).ToString();
                                ДвижениеДенФинОперПоступ.АкцДол = ДвижениеДенФинОперПоступАкцДол;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4314", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4314", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПоступОблВексДр = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4314", "Period1") != "") ДвижениеДенФинОперПоступОблВексДр.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4314", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4314", "Period2") != "") ДвижениеДенФинОперПоступОблВексДр.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4314", "Period2")).ToString();
                                ДвижениеДенФинОперПоступ.ОблВексДр = ДвижениеДенФинОперПоступОблВексДр;
                            }

                            //var ДвижениеДенФинОперПоступВПокФинПост = new ВПокОПТип(); //массив
                            //ДвижениеДенФинОперПоступВПокФинПост.НаимПок = "";
                            //ДвижениеДенФинОперПоступВПокФинПост.СумОтч = "";
                            //ДвижениеДенФинОперПоступВПокФинПост.СумПред = "";
                            if ((GetFields.gettabledata(Document, Processing, "4319", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4319", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПоступПрочПоступ = new ОП_ДТип();
                                if (GetFields.gettabledata(Document, Processing, "4319", "Period1") != "") ДвижениеДенФинОперПоступПрочПоступ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4319", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4319", "Period2") != "") ДвижениеДенФинОперПоступПрочПоступ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4319", "Period2")).ToString();
                                ДвижениеДенФинОперПоступ.ПрочПоступ = ДвижениеДенФинОперПоступПрочПоступ;
                            }
                            //var ДвижениеДенФинОперПоступПрочПоступВПокОП = new ВПокОПТип(); //массив
                            //ДвижениеДенФинОперПоступПрочПоступВПокОП.НаимПок = "";
                            //ДвижениеДенФинОперПоступПрочПоступВПокОП.СумОтч = "";
                            //ДвижениеДенФинОперПоступПрочПоступВПокОП.СумПред = "";

                            //Приравнивание
                            //ВПокОПТип[] _ДвижениеДенФинОперПоступПрочПоступВПокОП = { ДвижениеДенФинОперПоступПрочПоступВПокОП };
                            //ДвижениеДенФинОперПоступПрочПоступ.ВПокОП = _ДвижениеДенФинОперПоступПрочПоступВПокОП;

                            //ВПокОПТип[] _ДвижениеДенФинОперПоступВПокФинПост = { ДвижениеДенФинОперПоступВПокФинПост };
                            //ДвижениеДенФинОперПоступ.ВПокФинПост = _ДвижениеДенФинОперПоступВПокФинПост;

                            ДвижениеДенФинОпер.Поступ = ДвижениеДенФинОперПоступ;
                        }
                        #endregion //Конец ДвижениеДенТекОперПоступ

                        #region ДвижениеДенФинОперПлатеж
                        if ((GetFields.gettabledata(Document, Processing, "4320", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4320", "Period2") != ""))
                        {
                            var ДвижениеДенФинОперПлатеж = new ФайлДокументДвижениеДенФинОперПлатеж(); //описал вне региона
                            if (GetFields.gettabledata(Document, Processing, "4320", "Period1") != "") ДвижениеДенФинОперПлатеж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4320", "Period1")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "4320", "Period2") != "") ДвижениеДенФинОперПлатеж.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4320", "Period2")).ToString();

                            if ((GetFields.gettabledata(Document, Processing, "4321", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4321", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПлатежВыкупАкц = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4321", "Period1") != "") ДвижениеДенФинОперПлатежВыкупАкц.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4321", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4321", "Period2") != "") ДвижениеДенФинОперПлатежВыкупАкц.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4321", "Period2")).ToString();
                                ДвижениеДенФинОперПлатеж.ВыкупАкц = ДвижениеДенФинОперПлатежВыкупАкц;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4322", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4322", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПлатежУплДивИн = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4322", "Period1") != "") ДвижениеДенФинОперПлатежУплДивИн.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4322", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4322", "Period2") != "") ДвижениеДенФинОперПлатежУплДивИн.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4322", "Period2")).ToString();
                                ДвижениеДенФинОперПлатеж.УплДивИн = ДвижениеДенФинОперПлатежУплДивИн;
                            }
                            if ((GetFields.gettabledata(Document, Processing, "4323", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4323", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПлатежВыкВексКЗ = new ОПТип();
                                if (GetFields.gettabledata(Document, Processing, "4323", "Period1") != "") ДвижениеДенФинОперПлатежВыкВексКЗ.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4323", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4323", "Period2") != "") ДвижениеДенФинОперПлатежВыкВексКЗ.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4323", "Period2")).ToString();
                                ДвижениеДенФинОперПлатеж.ВыкВексКЗ = ДвижениеДенФинОперПлатежВыкВексКЗ;
                            }

                            //var ДвижениеДенФинОперПлатежВПокФинПлат = new ВПокОПТип(); //массив
                            //ДвижениеДенФинОперПлатежВПокФинПлат.НаимПок = "";
                            //ДвижениеДенФинОперПлатежВПокФинПлат.СумОтч = "";
                            //ДвижениеДенФинОперПлатежВПокФинПлат.СумПред = "";
                            if ((GetFields.gettabledata(Document, Processing, "4329", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4329", "Period2") != ""))
                            {
                                var ДвижениеДенФинОперПлатежПрочПлатеж = new ОП_ДТип();
                                if (GetFields.gettabledata(Document, Processing, "4329", "Period1") != "") ДвижениеДенФинОперПлатежПрочПлатеж.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4329", "Period1")).ToString();
                                if (GetFields.gettabledata(Document, Processing, "4329", "Period2") != "") ДвижениеДенФинОперПлатежПрочПлатеж.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4329", "Period2")).ToString();
                                ДвижениеДенФинОперПлатеж.ПрочПлатеж = ДвижениеДенФинОперПлатежПрочПлатеж;
                            }
                            //var ДвижениеДенФинОперПлатежПрочПлатежВПокОП = new ВПокОПТип(); //массив
                            //ДвижениеДенФинОперПлатежПрочПлатежВПокОП.НаимПок = "";
                            //ДвижениеДенФинОперПлатежПрочПлатежВПокОП.СумОтч = "";
                            //ДвижениеДенФинОперПлатежПрочПлатежВПокОП.СумПред = "";

                            //Приравнивание
                            //ВПокОПТип[] _ДвижениеДенФинОперПлатежПрочПлатежВПокОП = { ДвижениеДенФинОперПлатежПрочПлатежВПокОП };
                            //ДвижениеДенФинОперПлатежПрочПлатеж.ВПокОП = _ДвижениеДенФинОперПлатежПрочПлатежВПокОП;

                            //ВПокОПТип[] _ДвижениеДенФинОперПлатежВПокФинПлат = { ДвижениеДенФинОперПлатежВПокФинПлат };
                            //ДвижениеДенФинОперПлатеж.ВПокФинПлат = _ДвижениеДенФинОперПлатежВПокФинПлат;

                            ДвижениеДенФинОпер.Платеж = ДвижениеДенФинОперПлатеж;
                        }
                        #endregion //ДвижениеДенФинОперПлатеж

                        ДвижениеДенФинОпер.СальдоФин = ДвижениеДенФинОперСальдоФин;
                        ДвижениеДен.ФинОпер = ДвижениеДенФинОпер;
                    }
                    #endregion //конец третий группы 

                    if ((GetFields.gettabledata(Document, Processing, "4400", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4400", "Period2") != ""))
                    {
                        var ДвижениеДенСальдоОтч = new ОПТип();

                        if (GetFields.gettabledata(Document, Processing, "4400", "Period1") != "") ДвижениеДенСальдоОтч.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4400", "Period1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "4400", "Period2") != "") ДвижениеДенСальдоОтч.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4400", "Period2")).ToString();
                        ДвижениеДен.СальдоОтч = ДвижениеДенСальдоОтч;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "4450", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4450", "Period2") != ""))
                    {
                        var ДвижениеДенОстНачОтч = new ОПТип();
                        if (GetFields.gettabledata(Document, Processing, "4450", "Period1") != "") ДвижениеДенОстНачОтч.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4450", "Period1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "4450", "Period2") != "") ДвижениеДенОстНачОтч.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4450", "Period2")).ToString();
                        ДвижениеДен.ОстНачОтч = ДвижениеДенОстНачОтч;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "4500", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4500", "Period2") != ""))
                    {
                        var ДвижениеДенОстКонОтч = new ОПТип();
                        if (GetFields.gettabledata(Document, Processing, "4500", "Period1") != "") ДвижениеДенОстКонОтч.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4500", "Period1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "4500", "Period2") != "") ДвижениеДенОстКонОтч.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4500", "Period2")).ToString();
                        ДвижениеДен.ОстКонОтч = ДвижениеДенОстКонОтч;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "4490", "Period1") != "") || (GetFields.gettabledata(Document, Processing, "4490", "Period2") != ""))
                    {
                        var ДвижениеДенВлИзмКурс = new ОПТип();
                        if (GetFields.gettabledata(Document, Processing, "4490", "Period1") != "") ДвижениеДенВлИзмКурс.СумОтч = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4490", "Period1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "4490", "Period2") != "") ДвижениеДенВлИзмКурс.СумПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "4490", "Period2")).ToString();
                        ДвижениеДен.ВлИзмКурс = ДвижениеДенВлИзмКурс;
                    }

                    //Приравнивание

                    ДвижениеДен.ОКЕИ = ФайлДокументДвижениеДенОКЕИ.Item384;
                    ДвижениеДен.ОКУД = ДвижениеДенОКУД;
                    ДвижениеДен.Период = ДвижениеДенПериод;



                    ФайлДокументДвижениеДен[] _ДвижениеДен = { ДвижениеДен };
                    Документ.ДвижениеДен = _ДвижениеДен;
                }
                #endregion Документ_ДвижениеДен

                #region Документ_ОтчетИзмКап
                if (Document.DefinitionName == "3_140_Отчет об изменениях капитала")
                {
                    var ОтчетИзмКап = new ФайлДокументОтчетИзмКап();
                    var ОтчетИзмКапОКЕИ = new ФайлДокументОтчетИзмКапОКЕИ();
                    var ОтчетИзмКапОКУД = new ФайлДокументОтчетИзмКапОКУД();
                    var ОтчетИзмКапПериод = new ПериодГодТип();

                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Date")) && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out DateTime dateValue2)) { ОтчетИзмКап.ОтчетГод = dateValue2.Year.ToString(); /*Convert.ToDateTime(GetFields.GetField(Document, Processing, "Date")).Year.ToString();*/} else { ОтчетИзмКап.ОтчетГод = DateTime.Now.Year.ToString(); }


                    #region Документ_ОтчетИзмКапДвиженКап
                    //if ((GetFields.gettabledata(Document, Processing, "3100", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3100", "Itogo") != ""))
                    //{
                    var ОтчетИзмКапДвиженКап = new ФайлДокументОтчетИзмКапДвиженКап();

                    if ((GetFields.gettabledata(Document, Processing, "3100", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3100", "Itogo") != ""))
                    {
                        var ОтчетИзмКапДвиженКапКап31ДекПред = new ДвижКапПГод();
                        if (GetFields.gettabledata(Document, Processing, "3100", "Dobavochniy") != "") ОтчетИзмКапДвиженКапКап31ДекПред.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3100", "Dobavochniy")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3100", "Itogo") != "") ОтчетИзмКапДвиженКапКап31ДекПред.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3100", "Itogo")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3100", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапКап31ДекПред.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3100", "neraspredelenniy")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3100", "Rezervniy") != "") ОтчетИзмКапДвиженКапКап31ДекПред.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3100", "Rezervniy")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3100", "Sobstvennie") != "") ОтчетИзмКапДвиженКапКап31ДекПред.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3100", "Sobstvennie")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3100", "ustavnoy") != "") ОтчетИзмКапДвиженКапКап31ДекПред.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3100", "ustavnoy")).ToString();
                        ОтчетИзмКапДвиженКап.Кап31ДекПред = ОтчетИзмКапДвиженКапКап31ДекПред;
                    }
                    #region ОтчетИзмКапДвиженКапПредГод
                    if ((GetFields.gettabledata(Document, Processing, "3200", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3200", "Itogo") != ""))
                    {
                        var ОтчетИзмКапДвиженКапПредГод = new ДвижКапГодТип();

                        #region ОтчетИзмКапДвиженКапПредГодУвеличКапитал
                        var ОтчетИзмКапДвиженКапПредГодУвеличКапитал = new ДвижКапГодТипУвеличКапитал();

                        if ((GetFields.gettabledata(Document, Processing, "3210", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3210", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3210", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3210", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3210", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3210", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3210", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3210", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3210", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3210", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3210", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3210", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3210", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3210", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.УвеличКапВс = ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличКапВс;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3211", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталЧистПриб = new ДвижКапГодТипУвеличКапиталЧистПриб();
                            if (GetFields.gettabledata(Document, Processing, "3211", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталЧистПриб.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3211", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3211", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталЧистПриб.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3211", "Itogo")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.ЧистПриб = ОтчетИзмКапДвиженКапПредГодУвеличКапиталЧистПриб;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3212", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталПереоцИмущ = new ДвижКапГодТипУвеличКапиталПереоцИмущ();
                            if (GetFields.gettabledata(Document, Processing, "3212", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталПереоцИмущ.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3212", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3212", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталПереоцИмущ.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3212", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3212", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталПереоцИмущ.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3212", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.ПереоцИмущ = ОтчетИзмКапДвиженКапПредГодУвеличКапиталПереоцИмущ;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3213", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталДохУвелКап = new ДвижКапГодТипУвеличКапиталДохУвелКап();
                            if (GetFields.gettabledata(Document, Processing, "3213", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДохУвелКап.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3213", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3213", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДохУвелКап.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3213", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3213", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДохУвелКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3213", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.ДохУвелКап = ОтчетИзмКапДвиженКапПредГодУвеличКапиталДохУвелКап;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3214", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3214", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталДопВыпАкций = new ДвижКапГодТипУвеличКапиталДопВыпАкций();
                            if (GetFields.gettabledata(Document, Processing, "3214", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДопВыпАкций.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3214", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3214", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДопВыпАкций.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3214", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3214", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДопВыпАкций.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3214", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3214", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталДопВыпАкций.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3214", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.ДопВыпАкций = ОтчетИзмКапДвиженКапПредГодУвеличКапиталДопВыпАкций;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3215", "ustavnoy") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличНомАкц = new ДвижКапГодТипУвеличКапиталУвеличНомАкц();
                            if (GetFields.gettabledata(Document, Processing, "3215", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличНомАкц.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3215", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3215", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличНомАкц.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3215", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3215", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличНомАкц.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3215", "ustavnoy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3215", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличНомАкц.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3215", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.УвеличНомАкц = ОтчетИзмКапДвиженКапПредГодУвеличКапиталУвеличНомАкц;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3216", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3216", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3216", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3216", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3216", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3216", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3216", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3216", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3216", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3216", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3216", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3216", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3216", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3216", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУвеличКапитал.Реорганизация = ОтчетИзмКапДвиженКапПредГодУвеличКапиталРеорганизация;
                        }
                        //var ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап = new ВПокДвижКапПГод();
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.ДобКапитал = "";
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.Итог = "";
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.НаимПок = "";
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.НераспПриб = "";
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.РезКапитал = "";
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.СобВыкупАкц = "";
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап.УстКапитал = "";

                        //ВПокДвижКапПГод[] items = { ОтчетИзмКапДвиженКапПредГодУвеличКапиталВПокУвелКап };
                        //ОтчетИзмКапДвиженКапПредГодУвеличКапитал.ВПокУвелКап = items;

                        ОтчетИзмКапДвиженКапПредГод.УвеличКапитал = ОтчетИзмКапДвиженКапПредГодУвеличКапитал;
                        #endregion


                        #region ОтчетИзмКапДвиженКапПредГодУменКапитал

                        var ОтчетИзмКапДвиженКапПредГодУменКапитал = new ДвижКапГодТипУменКапитал();

                        if ((GetFields.gettabledata(Document, Processing, "3220", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3220", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3220", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3220", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3220", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3220", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3220", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3220", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3220", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3220", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3220", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3220", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3220", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3220", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.УменКапВс = ОтчетИзмКапДвиженКапПредГодУменКапиталУменКапВс;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3221", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталУбыток = new ДвижКапГодТипУменКапиталУбыток();
                            if (GetFields.gettabledata(Document, Processing, "3221", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУбыток.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3221", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3221", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУбыток.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3221", "Itogo")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.Убыток = ОтчетИзмКапДвиженКапПредГодУменКапиталУбыток;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3222", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталПереоцИмущ = new ДвижКапГодТипУменКапиталПереоцИмущ();

                            if (GetFields.gettabledata(Document, Processing, "3222", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталПереоцИмущ.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3222", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3222", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталПереоцИмущ.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3222", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3222", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталПереоцИмущ.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3222", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.ПереоцИмущ = ОтчетИзмКапДвиженКапПредГодУменКапиталПереоцИмущ;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3223", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталРасхУменКап = new ДвижКапГодТипУменКапиталРасхУменКап();
                            if (GetFields.gettabledata(Document, Processing, "3223", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРасхУменКап.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3223", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3223", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРасхУменКап.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3223", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3223", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРасхУменКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3223", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.РасхУменКап = ОтчетИзмКапДвиженКапПредГодУменКапиталРасхУменКап;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3224", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3224", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц = new ДвижКапГодТипУменКапиталУменНомАкц();
                            if (GetFields.gettabledata(Document, Processing, "3224", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3224", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3224", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3224", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3224", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3224", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3224", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3224", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3224", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3224", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.УменНомАкц = ОтчетИзмКапДвиженКапПредГодУменКапиталУменНомАкц;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3225", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3224", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций = new ДвижКапГодТипУменКапиталУменКолАкций();
                            if (GetFields.gettabledata(Document, Processing, "3225", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3225", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3225", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3225", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3225", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3225", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3225", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3225", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3225", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3225", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.УменКолАкций = ОтчетИзмКапДвиженКапПредГодУменКапиталУменКолАкций;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3226", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3226", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3226", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3226", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3226", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3226", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3226", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3226", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3226", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3226", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3226", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3226", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3226", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3226", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.Реорганизация = ОтчетИзмКапДвиженКапПредГодУменКапиталРеорганизация;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3227", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапПредГодУменКапиталДивиденды = new ДвижКапГодТипУменКапиталДивиденды();
                            ОтчетИзмКапДвиженКапПредГодУменКапиталДивиденды.Итог = "";
                            ОтчетИзмКапДвиженКапПредГодУменКапиталДивиденды.НераспПриб = "";
                            if (GetFields.gettabledata(Document, Processing, "3227", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталДивиденды.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3227", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3227", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодУменКапиталДивиденды.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3227", "Itogo")).ToString();
                            ОтчетИзмКапДвиженКапПредГодУменКапитал.Дивиденды = ОтчетИзмКапДвиженКапПредГодУменКапиталДивиденды;
                        }

                        ОтчетИзмКапДвиженКапПредГод.УменКапитал = ОтчетИзмКапДвиженКапПредГодУменКапитал;
                        #endregion

                        if ((GetFields.gettabledata(Document, Processing, "3230", "Dobavochniy") != "" || GetFields.gettabledata(Document, Processing, "3230", "Rezervniy") != "" || GetFields.gettabledata(Document, Processing, "3230", "neraspredelenniy") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодИзмДобавКап = new ДвижКапГодТипИзмДобавКап();
                            if (GetFields.gettabledata(Document, Processing, "3230", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодИзмДобавКап.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3230", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3230", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодИзмДобавКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3230", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3230", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодИзмДобавКап.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3230", "Rezervniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГод.ИзмДобавКап = ОтчетИзмКапДвиженКапПредГодИзмДобавКап;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3240", "Rezervniy") != "" || GetFields.gettabledata(Document, Processing, "3240", "neraspredelenniy") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодИзмРезервКап = new ДвижКапГодТипИзмРезервКап();
                            ОтчетИзмКапДвиженКапПредГодИзмРезервКап.НераспПриб = "";
                            ОтчетИзмКапДвиженКапПредГодИзмРезервКап.РезКапитал = "";
                            if (GetFields.gettabledata(Document, Processing, "3240", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодИзмРезервКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3240", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3240", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодИзмРезервКап.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3240", "Rezervniy")).ToString();
                            ОтчетИзмКапДвиженКапПредГод.ИзмРезервКап = ОтчетИзмКапДвиженКапПредГодИзмРезервКап;
                        }
                        //var ОтчетИзмКапДвиженКапПредГодВПокДвижКап = new ВПокДвижКапПГод();
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.ДобКапитал = "";
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.Итог = "";
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.НаимПок = "";
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.НераспПриб = "";
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.РезКапитал = "";
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.СобВыкупАкц = "";
                        //ОтчетИзмКапДвиженКапПредГодВПокДвижКап.УстКапитал = "";

                        //ВПокДвижКапПГод[] items1 = { ОтчетИзмКапДвиженКапПредГодВПокДвижКап };
                        //ОтчетИзмКапДвиженКапПредГод.ВПокДвижКап = items1;
                        if ((GetFields.gettabledata(Document, Processing, "3200", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3200", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапПредГодКап31дек = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3200", "Dobavochniy") != "") ОтчетИзмКапДвиженКапПредГодКап31дек.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3200", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3200", "Itogo") != "") ОтчетИзмКапДвиженКапПредГодКап31дек.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3200", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3200", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапПредГодКап31дек.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3200", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3200", "Rezervniy") != "") ОтчетИзмКапДвиженКапПредГодКап31дек.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3200", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3200", "Sobstvennie") != "") ОтчетИзмКапДвиженКапПредГодКап31дек.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3200", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3200", "ustavnoy") != "") ОтчетИзмКапДвиженКапПредГодКап31дек.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3200", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапПредГод.Кап31дек = ОтчетИзмКапДвиженКапПредГодКап31дек;
                        }

                        ОтчетИзмКапДвиженКап.ПредГод = ОтчетИзмКапДвиженКапПредГод;
                    }
                    #endregion
                    #region ОтчетИзмКапДвиженКапОтчетГод
                    if ((GetFields.gettabledata(Document, Processing, "3300", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3300", "Itogo") != ""))
                    {

                        var ОтчетИзмКапДвиженКапОтчетГод = new ДвижКапГодТип();

                        #region ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал
                        var ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал = new ДвижКапГодТипУвеличКапитал();
                        if ((GetFields.gettabledata(Document, Processing, "3310", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3310", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3310", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3310", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3310", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3310", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3310", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3310", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3310", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3310", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3310", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3310", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3310", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3310", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.УвеличКапВс = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличКапВс;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3311", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталЧистПриб = new ДвижКапГодТипУвеличКапиталЧистПриб();
                            if (GetFields.gettabledata(Document, Processing, "3311", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталЧистПриб.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3311", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3311", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталЧистПриб.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3311", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.ЧистПриб = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталЧистПриб;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3312", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталПереоцИмущ = new ДвижКапГодТипУвеличКапиталПереоцИмущ();
                            if (GetFields.gettabledata(Document, Processing, "3312", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталПереоцИмущ.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3312", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3312", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталПереоцИмущ.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3312", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3312", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталПереоцИмущ.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3312", "Dobavochniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.ПереоцИмущ = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталПереоцИмущ;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3313", "Itogo") != "")
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДохУвелКап = new ДвижКапГодТипУвеличКапиталДохУвелКап();
                            if (GetFields.gettabledata(Document, Processing, "3313", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДохУвелКап.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3313", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3313", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДохУвелКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3313", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3313", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДохУвелКап.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3313", "Dobavochniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.ДохУвелКап = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДохУвелКап;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3314", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3314", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДопВыпАкций = new ДвижКапГодТипУвеличКапиталДопВыпАкций();
                            if (GetFields.gettabledata(Document, Processing, "3314", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДопВыпАкций.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3314", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3314", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДопВыпАкций.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3314", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3314", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДопВыпАкций.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3314", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3314", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДопВыпАкций.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3314", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.ДопВыпАкций = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталДопВыпАкций;
                        }
                        if (GetFields.gettabledata(Document, Processing, "3315", "ustavnoy") != "")
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличНомАкц = new ДвижКапГодТипУвеличКапиталУвеличНомАкц();
                            if (GetFields.gettabledata(Document, Processing, "3315", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличНомАкц.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3315", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3315", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличНомАкц.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3315", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3315", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличНомАкц.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3315", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3315", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличНомАкц.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3315", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.УвеличНомАкц = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталУвеличНомАкц;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3316", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3316", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3316", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3316", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3316", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3316", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3316", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3316", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3316", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3316", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3316", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3316", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3316", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3316", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.Реорганизация = ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталРеорганизация;
                        }
                        //var ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап = new ВПокДвижКапПГод();
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.ДобКапитал = "";
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.Итог = "";
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.НаимПок = "";
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.НераспПриб = "";
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.РезКапитал = "";
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.СобВыкупАкц = "";
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап.УстКапитал = "";

                        //ВПокДвижКапПГод[] items2 = { ОтчетИзмКапДвиженКапОтчетГодУвеличКапиталВПокУвелКап };
                        //ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал.ВПокУвелКап = items2;

                        ОтчетИзмКапДвиженКапОтчетГод.УвеличКапитал = ОтчетИзмКапДвиженКапОтчетГодУвеличКапитал;
                        #endregion


                        #region ОтчетИзмКапДвиженКапОтчетГодУменКапитал
                        var ОтчетИзмКапДвиженКапОтчетГодУменКапитал = new ДвижКапГодТипУменКапитал();
                        if ((GetFields.gettabledata(Document, Processing, "3320", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3320", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3320", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3320", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3320", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3320", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3320", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3320", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3320", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3320", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3320", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3320", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3320", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3320", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.УменКапВс = ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКапВс;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3321", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталУбыток = new ДвижКапГодТипУменКапиталУбыток();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапиталУбыток.Итог = "";
                            ОтчетИзмКапДвиженКапОтчетГодУменКапиталУбыток.НераспПриб = "";
                            if (GetFields.gettabledata(Document, Processing, "3321", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУбыток.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3321", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3321", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУбыток.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3321", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.Убыток = ОтчетИзмКапДвиженКапОтчетГодУменКапиталУбыток;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3322", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталПереоцИмущ = new ДвижКапГодТипУменКапиталПереоцИмущ();
                            if (GetFields.gettabledata(Document, Processing, "3322", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталПереоцИмущ.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3322", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3322", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталПереоцИмущ.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3322", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3322", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталПереоцИмущ.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3322", "Dobavochniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.ПереоцИмущ = ОтчетИзмКапДвиженКапОтчетГодУменКапиталПереоцИмущ;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3323", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталРасхУменКап = new ДвижКапГодТипУменКапиталРасхУменКап();
                            if (GetFields.gettabledata(Document, Processing, "3323", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРасхУменКап.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3323", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3323", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРасхУменКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3323", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3323", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРасхУменКап.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3323", "Dobavochniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.РасхУменКап = ОтчетИзмКапДвиженКапОтчетГодУменКапиталРасхУменКап;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3324", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3324", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц = new ДвижКапГодТипУменКапиталУменНомАкц();
                            if (GetFields.gettabledata(Document, Processing, "3324", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3324", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3324", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3324", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3324", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3324", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3324", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3324", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3324", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3324", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.УменНомАкц = ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменНомАкц;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3325", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3325", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций = new ДвижКапГодТипУменКапиталУменКолАкций();
                            if (GetFields.gettabledata(Document, Processing, "3325", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3325", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3325", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3325", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3325", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3325", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3325", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3325", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3325", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3325", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.УменКолАкций = ОтчетИзмКапДвиженКапОтчетГодУменКапиталУменКолАкций;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3326", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3326", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3326", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3326", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3326", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3326", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3326", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3326", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3326", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3326", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3326", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3326", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3326", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3326", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.Реорганизация = ОтчетИзмКапДвиженКапОтчетГодУменКапиталРеорганизация;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3327", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодУменКапиталДивиденды = new ДвижКапГодТипУменКапиталДивиденды();
                            if (GetFields.gettabledata(Document, Processing, "3327", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталДивиденды.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3327", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3327", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодУменКапиталДивиденды.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3327", "neraspredelenniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГодУменКапитал.Дивиденды = ОтчетИзмКапДвиженКапОтчетГодУменКапиталДивиденды;
                        }
                        ОтчетИзмКапДвиженКапОтчетГод.УменКапитал = ОтчетИзмКапДвиженКапОтчетГодУменКапитал;
                        #endregion

                        if ((GetFields.gettabledata(Document, Processing, "3330", "Dobavochniy") != "" || GetFields.gettabledata(Document, Processing, "3330", "Rezervniy") != "" || GetFields.gettabledata(Document, Processing, "3330", "neraspredelenniy") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап = new ДвижКапГодТипИзмДобавКап();
                            ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап.ДобКапитал = "";
                            ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап.НераспПриб = "";
                            ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап.РезКапитал = "";
                            if (GetFields.gettabledata(Document, Processing, "3330", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3330", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3330", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3330", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3330", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3330", "Rezervniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГод.ИзмДобавКап = ОтчетИзмКапДвиженКапОтчетГодИзмДобавКап;
                        }
                        if ((GetFields.gettabledata(Document, Processing, "3340", "Rezervniy") != "" || GetFields.gettabledata(Document, Processing, "3340", "neraspredelenniy") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодИзмРезервКап = new ДвижКапГодТипИзмРезервКап();

                            if (GetFields.gettabledata(Document, Processing, "3330", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодИзмРезервКап.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3330", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3330", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодИзмРезервКап.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3330", "Rezervniy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГод.ИзмРезервКап = ОтчетИзмКапДвиженКапОтчетГодИзмРезервКап;
                        }

                        //    var ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап = new ВПокДвижКапПГод();
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.ДобКапитал = "";
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.Итог = "";
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.НаимПок = "";
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.НераспПриб = "";
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.РезКапитал = "";
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.СобВыкупАкц = "";
                        //ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап.УстКапитал = "";

                        //ВПокДвижКапПГод[] items3 = { ОтчетИзмКапДвиженКапОтчетГодВПокДвижКап };
                        //ОтчетИзмКапДвиженКапОтчетГод.ВПокДвижКап = items3;
                        if ((GetFields.gettabledata(Document, Processing, "3330", "ustavnoy") != "") || (GetFields.gettabledata(Document, Processing, "3300", "Itogo") != ""))
                        {
                            var ОтчетИзмКапДвиженКапОтчетГодКап31дек = new ДвижКапПГод();
                            if (GetFields.gettabledata(Document, Processing, "3300", "Dobavochniy") != "") ОтчетИзмКапДвиженКапОтчетГодКап31дек.ДобКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3300", "Dobavochniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3300", "Itogo") != "") ОтчетИзмКапДвиженКапОтчетГодКап31дек.Итог = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3300", "Itogo")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3300", "neraspredelenniy") != "") ОтчетИзмКапДвиженКапОтчетГодКап31дек.НераспПриб = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3300", "neraspredelenniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3300", "Rezervniy") != "") ОтчетИзмКапДвиженКапОтчетГодКап31дек.РезКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3300", "Rezervniy")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3300", "Sobstvennie") != "") ОтчетИзмКапДвиженКапОтчетГодКап31дек.СобВыкупАкц = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3300", "Sobstvennie")).ToString();
                            if (GetFields.gettabledata(Document, Processing, "3300", "ustavnoy") != "") ОтчетИзмКапДвиженКапОтчетГодКап31дек.УстКапитал = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3300", "ustavnoy")).ToString();
                            ОтчетИзмКапДвиженКапОтчетГод.Кап31дек = ОтчетИзмКапДвиженКапОтчетГодКап31дек;
                        }
                        ОтчетИзмКапДвиженКап.ОтчетГод = ОтчетИзмКапДвиженКапОтчетГод;
                    }
                    #endregion

                    ОтчетИзмКап.ДвиженКап = ОтчетИзмКапДвиженКап;
                    #endregion



                    #region ОтчетИзмКапКоррект

                    var ОтчетИзмКапКоррект = new ФайлДокументОтчетИзмКапКоррект();

                    #region ОтчетИзмКапКорректКапитВС 
                    var ОтчетИзмКапКорректКапитВС = new КорКапТип();
                    if ((GetFields.gettabledata(Document, Processing, "3400", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3400", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3400", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректКапитВСДоКоррект = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3400", "Na1") != "") ОтчетИзмКапКорректКапитВСДоКоррект.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3400", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3400", "Za1") != "") ОтчетИзмКапКорректКапитВСДоКоррект.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3400", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3400", "Za2") != "") ОтчетИзмКапКорректКапитВСДоКоррект.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3400", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3400", "Na2") != "") ОтчетИзмКапКорректКапитВСДоКоррект.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3400", "Na2")).ToString();
                        ОтчетИзмКапКорректКапитВС.ДоКоррект = ОтчетИзмКапКорректКапитВСДоКоррект;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3410", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3410", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3410", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректКапитВСКорИзмУчПол = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3410", "Na1") != "") ОтчетИзмКапКорректКапитВСКорИзмУчПол.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3410", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3410", "Za1") != "") ОтчетИзмКапКорректКапитВСКорИзмУчПол.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3410", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3410", "Za2") != "") ОтчетИзмКапКорректКапитВСКорИзмУчПол.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3410", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3410", "Na2") != "") ОтчетИзмКапКорректКапитВСКорИзмУчПол.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3410", "Na2")).ToString();
                        ОтчетИзмКапКорректКапитВС.КорИзмУчПол = ОтчетИзмКапКорректКапитВСКорИзмУчПол;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3420", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3420", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3420", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректКапитВСКорИспрОш = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3420", "Na1") != "") ОтчетИзмКапКорректКапитВСКорИспрОш.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3420", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3420", "Za1") != "") ОтчетИзмКапКорректКапитВСКорИспрОш.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3420", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3420", "Za2") != "") ОтчетИзмКапКорректКапитВСКорИспрОш.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3420", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3420", "Na2") != "") ОтчетИзмКапКорректКапитВСКорИспрОш.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3420", "Na2")).ToString();
                        ОтчетИзмКапКорректКапитВС.КорИспрОш = ОтчетИзмКапКорректКапитВСКорИспрОш;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3500", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3500", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3500", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректКапитВСПослеКоррект = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3500", "Na1") != "") ОтчетИзмКапКорректКапитВСПослеКоррект.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3500", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3500", "Za1") != "") ОтчетИзмКапКорректКапитВСПослеКоррект.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3500", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3500", "Za2") != "") ОтчетИзмКапКорректКапитВСПослеКоррект.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3500", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3500", "Na2") != "") ОтчетИзмКапКорректКапитВСПослеКоррект.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3500", "Na2")).ToString();
                        ОтчетИзмКапКорректКапитВС.ПослеКоррект = ОтчетИзмКапКорректКапитВСПослеКоррект;
                    }
                    ОтчетИзмКапКоррект.КапитВс = ОтчетИзмКапКорректКапитВС;
                    #endregion //ОтчетИзмКапКорректКапитВС 


                    #region ОтчетИзмКапКорректНераспПриб
                    var ОтчетИзмКапКорректНераспПриб = new КорКапТип();
                    if ((GetFields.gettabledata(Document, Processing, "3401", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3401", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3401", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректНераспПрибДоКоррект = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3401", "Na1") != "") ОтчетИзмКапКорректНераспПрибДоКоррект.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3401", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3401", "Za1") != "") ОтчетИзмКапКорректНераспПрибДоКоррект.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3401", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3401", "Za2") != "") ОтчетИзмКапКорректНераспПрибДоКоррект.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3401", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3401", "Na2") != "") ОтчетИзмКапКорректНераспПрибДоКоррект.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3401", "Na2")).ToString();
                        ОтчетИзмКапКорректНераспПриб.ДоКоррект = ОтчетИзмКапКорректНераспПрибДоКоррект;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3411", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3411", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3411", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректНераспПрибКорИзмУчПол = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3411", "Na1") != "") ОтчетИзмКапКорректНераспПрибКорИзмУчПол.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3411", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3411", "Za1") != "") ОтчетИзмКапКорректНераспПрибКорИзмУчПол.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3411", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3411", "Za2") != "") ОтчетИзмКапКорректНераспПрибКорИзмУчПол.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3411", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3411", "Na2") != "") ОтчетИзмКапКорректНераспПрибКорИзмУчПол.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3411", "Na2")).ToString();
                        ОтчетИзмКапКорректНераспПриб.КорИзмУчПол = ОтчетИзмКапКорректНераспПрибКорИзмУчПол;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3421", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3421", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3421", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректНераспПрибКорИспрОш = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3421", "Na1") != "") ОтчетИзмКапКорректНераспПрибКорИспрОш.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3421", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3421", "Za1") != "") ОтчетИзмКапКорректНераспПрибКорИспрОш.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3421", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3421", "Za2") != "") ОтчетИзмКапКорректНераспПрибКорИспрОш.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3421", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3421", "Na2") != "") ОтчетИзмКапКорректНераспПрибКорИспрОш.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3421", "Na2")).ToString();
                        ОтчетИзмКапКорректНераспПриб.КорИспрОш = ОтчетИзмКапКорректНераспПрибКорИспрОш;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3501", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3501", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3501", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректНераспПрибПослеКоррект = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3501", "Na1") != "") ОтчетИзмКапКорректНераспПрибПослеКоррект.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3501", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3501", "Za1") != "") ОтчетИзмКапКорректНераспПрибПослеКоррект.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3501", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3501", "Za2") != "") ОтчетИзмКапКорректНераспПрибПослеКоррект.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3501", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3501", "Na2") != "") ОтчетИзмКапКорректНераспПрибПослеКоррект.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3501", "Na2")).ToString();
                        ОтчетИзмКапКорректНераспПриб.ПослеКоррект = ОтчетИзмКапКорректНераспПрибПослеКоррект;
                    }
                    ОтчетИзмКапКоррект.НераспПриб = ОтчетИзмКапКорректНераспПриб;
                    #endregion //ОтчетИзмКапКорректНераспПриб 

                    #region ОтчетИзмКапКорректДрСтатКап 
                    var ОтчетИзмКапКорректДрСтатКап = new КорКапТип();
                    // if ((GetFields.gettabledata(Document, Processing, "3402", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3402", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3402", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректДрСтатКапДоКоррект = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3402", "Na1") != "") { ОтчетИзмКапКорректДрСтатКапДоКоррект.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3402", "Na1")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапДоКоррект.На31ДекПрПред = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "3402", "Za1") != "") { ОтчетИзмКапКорректДрСтатКапДоКоррект.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3402", "Za1")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапДоКоррект.ИзмКапЧистПр = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "3402", "Za2") != "") { ОтчетИзмКапКорректДрСтатКапДоКоррект.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3402", "Za2")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапДоКоррект.ИзмКапИнФакт = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "3402", "Na2") != "") { ОтчетИзмКапКорректДрСтатКапДоКоррект.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3402", "Na2")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапДоКоррект.На31ДекПред = "0"; }
                        ОтчетИзмКапКорректДрСтатКап.ДоКоррект = ОтчетИзмКапКорректДрСтатКапДоКоррект;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3412", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3412", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3412", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректДрСтатКапКорИзмУчПол = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3412", "Na1") != "") ОтчетИзмКапКорректДрСтатКапКорИзмУчПол.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3412", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3412", "Za1") != "") ОтчетИзмКапКорректДрСтатКапКорИзмУчПол.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3412", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3412", "Za2") != "") ОтчетИзмКапКорректДрСтатКапКорИзмУчПол.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3412", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3412", "Na2") != "") ОтчетИзмКапКорректДрСтатКапКорИзмУчПол.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3412", "Na2")).ToString();
                        ОтчетИзмКапКорректДрСтатКап.КорИзмУчПол = ОтчетИзмКапКорректДрСтатКапКорИзмУчПол;
                    }
                    if ((GetFields.gettabledata(Document, Processing, "3422", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3422", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3422", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректДрСтатКапКорИспрОш = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3422", "Na1") != "") ОтчетИзмКапКорректДрСтатКапКорИспрОш.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3422", "Na1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3422", "Za1") != "") ОтчетИзмКапКорректДрСтатКапКорИспрОш.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3422", "Za1")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3422", "Za2") != "") ОтчетИзмКапКорректДрСтатКапКорИспрОш.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3422", "Za2")).ToString();
                        if (GetFields.gettabledata(Document, Processing, "3422", "Na2") != "") ОтчетИзмКапКорректДрСтатКапКорИспрОш.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3422", "Na2")).ToString();
                        ОтчетИзмКапКорректДрСтатКап.КорИспрОш = ОтчетИзмКапКорректДрСтатКапКорИспрОш;
                    }
                    // if ((GetFields.gettabledata(Document, Processing, "3502", "Za1") != "" || GetFields.gettabledata(Document, Processing, "3502", "Za2") != "" || GetFields.gettabledata(Document, Processing, "3502", "Na2") != ""))
                    {
                        var ОтчетИзмКапКорректДрСтатКапПослеКоррект = new КорКапПрТип();
                        if (GetFields.gettabledata(Document, Processing, "3502", "Na1") != "") { ОтчетИзмКапКорректДрСтатКапПослеКоррект.На31ДекПрПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3502", "Na1")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапПослеКоррект.На31ДекПрПред = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "3502", "Za1") != "") { ОтчетИзмКапКорректДрСтатКапПослеКоррект.ИзмКапЧистПр = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3502", "Za1")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапПослеКоррект.ИзмКапЧистПр = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "3502", "Za2") != "") { ОтчетИзмКапКорректДрСтатКапПослеКоррект.ИзмКапИнФакт = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3502", "Za2")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапПослеКоррект.ИзмКапИнФакт = "0"; }
                        if (GetFields.gettabledata(Document, Processing, "3502", "Na2") != "") { ОтчетИзмКапКорректДрСтатКапПослеКоррект.На31ДекПред = Convert.ToInt32(GetFields.gettabledata(Document, Processing, "3502", "Na2")).ToString(); } else { ОтчетИзмКапКорректДрСтатКапПослеКоррект.На31ДекПред = "0"; }
                        ОтчетИзмКапКорректДрСтатКап.ПослеКоррект = ОтчетИзмКапКорректДрСтатКапПослеКоррект;
                    }
                    ОтчетИзмКапКоррект.ДрСтатКап = ОтчетИзмКапКорректДрСтатКап;
                    #endregion //ОтчетИзмКапКорректДрСтатКап 





                    ОтчетИзмКап.Коррект = ОтчетИзмКапКоррект;
                    #endregion //ОтчетИзмКапКоррект

                    if (GetFields.GetField(Document, Processing, "ChAPer1") != "" || GetFields.GetField(Document, Processing, "ChAPer2") != "" || GetFields.GetField(Document, Processing, "ChAPer3") != "")
                    {
                        var ОтчетИзмКапЧистАктив = new ФайлДокументОтчетИзмКапЧистАктив();
                        if (GetFields.GetField(Document, Processing, "ChAPer1") != "") ОтчетИзмКапЧистАктив.На31ДекОтч = Convert.ToInt32(GetFields.GetField(Document, Processing, "ChAPer1")).ToString();
                        if (GetFields.GetField(Document, Processing, "ChAPer2") != "") ОтчетИзмКапЧистАктив.На31ДекПред = Convert.ToInt32(GetFields.GetField(Document, Processing, "ChAPer2")).ToString();
                        if (GetFields.GetField(Document, Processing, "ChAPer3") != "") ОтчетИзмКапЧистАктив.На31ДекПрПред = Convert.ToInt32(GetFields.GetField(Document, Processing, "ChAPer3")).ToString();
                        ОтчетИзмКап.ЧистАктив = ОтчетИзмКапЧистАктив;
                    }

                    ОтчетИзмКап.ОКЕИ = ФайлДокументОтчетИзмКапОКЕИ.Item384;
                    ОтчетИзмКап.ОКУД = ОтчетИзмКапОКУД;
                    ОтчетИзмКап.Период = ОтчетИзмКапПериод;

                    ФайлДокументОтчетИзмКап[] _ОтчетИзмКап = { ОтчетИзмКап };
                    Документ.ОтчетИзмКап = _ОтчетИзмКап;
                }
                #endregion //ОтчетИзмКап

                #region Документ_целИсп
                if (Document.DefinitionName == "3_341_Отчет о целевом использовании средств")
                {
                    var ЦелИсп = new ФайлДокументЦелИсп();
                    var ЦелИспОКЕИ = new ФайлДокументЦелИспОКЕИ();
                    if (GetFields.GetField(Document, Processing, "Date") != "" && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out DateTime dateValue3)) { ЦелИсп.ОтчетГод = dateValue3.Year.ToString(); /*Convert.ToDateTime(GetFields.GetField(Document, Processing, "Date")).Year.ToString();*/} else { ЦелИсп.ОтчетГод = DateTime.Now.Year.ToString(); }
                    //if (GetFields.GetField(Document, Processing, "Date") != "") { ЦелИсп.ОтчетГод = GetFields.GetField(Document, Processing, "Date"); } else { ЦелИсп.ОтчетГод = DateTime.Now.Year.ToString(); }//;
                    ЦелИсп.Период = ПериодГодТип.Item34;
                    var ЦелИспОКУД = new ФайлДокументЦелИспОКУД();

                    #region Поступило
                    if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period1")) || !string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period2")))
                    {
                        var Поступило = new ФайлДокументЦелИспПоступило();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period1"))) { Поступило.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period1"); } else { Поступило.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period2"))) { Поступило.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period2"); } else { Поступило.СумПред = "0"; }

                        var ПоступилоВступВзнос = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period1"))) { ПоступилоВступВзнос.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6210", "Kod", "Period1"); } else { ПоступилоВступВзнос.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6200", "Kod", "Period1"))) { ПоступилоВступВзнос.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6210", "Kod", "Period2"); } else { ПоступилоВступВзнос.СумПред = "0"; }

                        var ЧленВзнос = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6215", "Kod", "Period1"))) { ЧленВзнос.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6215", "Kod", "Period1"); } else { ЧленВзнос.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6215", "Kod", "Period2"))) { ЧленВзнос.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6215", "Kod", "Period2"); } else { ЧленВзнос.СумПред = "0"; }

                        var ЦелевВзнос = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6220", "Kod", "Period1"))) { ЦелевВзнос.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6220", "Kod", "Period1"); } else { ЦелевВзнос.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6220", "Kod", "Period2"))) { ЦелевВзнос.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6220", "Kod", "Period2"); } else { ЦелевВзнос.СумПред = "0"; }

                        var ДобрИмВзнос = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6230", "Kod", "Period1"))) { ДобрИмВзнос.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6230", "Kod", "Period1"); } else { ДобрИмВзнос.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6230", "Kod", "Period2"))) { ДобрИмВзнос.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6230", "Kod", "Period2"); } else { ДобрИмВзнос.СумПред = "0"; }

                        var ПрибПредДеят = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6240", "Kod", "Period1"))) { ПрибПредДеят.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6240", "Kod", "Period1"); } else { ПрибПредДеят.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6240", "Kod", "Period2"))) { ПрибПредДеят.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6240", "Kod", "Period2"); } else { ПрибПредДеят.СумПред = "0"; }



                        //var ВПокПоступ = new ВПокОПТип();
                        //ВПокПоступ.СумОтч = " ";
                        //ВПокПоступ.НаимПок = " ";
                        //ВПокПоступ.СумПред = " ";

                        #region Прочие
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6250", "Kod", "Period1")) || !string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6250", "Kod", "Period2")))
                        {
                            var Прочие = new ОП_ДТип();
                            if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6250", "Kod", "Period1"))) { Прочие.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6250", "Kod", "Period1"); } else { Прочие.СумОтч = "0"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6250", "Kod", "Period2"))) { Прочие.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6250", "Kod", "Period1"); } else { Прочие.СумПред = "0"; }

                            //var ВПокОП = new ВПокОПТип();
                            //ВПокОП.СумОтч = " ";
                            //ВПокОП.НаимПок = " ";
                            //ВПокОП.СумПред = " ";

                            Поступило.Прочие = Прочие;
                        }


                        #endregion Прочие


                        Поступило.ВступВзнос = ПоступилоВступВзнос;
                        Поступило.ЧленВзнос = ЧленВзнос;
                        Поступило.ДобрИмВзнос = ДобрИмВзнос;
                        Поступило.ЦелевВзнос = ЦелевВзнос;
                        Поступило.ПрибПредДеят = ПрибПредДеят;
                        ЦелИсп.Поступило = Поступило;
                    }
                    #endregion

                    #region Использовано
                    if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6300", "Kod", "Period1")) || !string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6300", "Kod", "Period2")))
                    {
                        var Использовано = new ФайлДокументЦелИспИспользовано();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6240", "Kod", "Period1"))) { Использовано.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6240", "Kod", "Period1"); } else { Использовано.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6300", "Kod", "Period2"))) { Использовано.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6300", "Kod", "Period2"); } else { Использовано.СумПред = "0"; }

                        #region РасхЦелМер
                        var РасхЦелМер = new ФайлДокументЦелИспИспользованоРасхЦелМер();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6310", "Kod", "Period1"))) { РасхЦелМер.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6310", "Kod", "Period1"); } else { РасхЦелМер.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6310", "Kod", "Period2"))) { РасхЦелМер.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6310", "Kod", "Period2"); } else { РасхЦелМер.СумПред = "0"; }

                        var СоцПом = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6311", "Kod", "Period1"))) { СоцПом.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6311", "Kod", "Period1"); } else { СоцПом.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6311", "Kod", "Period2"))) { СоцПом.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6311", "Kod", "Period2"); } else { СоцПом.СумПред = "0"; }

                        var ПровСемин = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6312", "Kod", "Period1"))) { ПровСемин.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6312", "Kod", "Period1"); } else { ПровСемин.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6312", "Kod", "Period2"))) { ПровСемин.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6312", "Kod", "Period2"); } else { ПровСемин.СумПред = "0"; }


                        //var ВПокЦелМер = new ВПокОПТип();
                        //ВПокЦелМер.СумОтч = " ";
                        //ВПокЦелМер.НаимПок = " ";
                        //ВПокЦелМер.СумПред = " ";

                        var ИныеМер = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6313", "Kod", "Period1"))) { ИныеМер.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6313", "Kod", "Period1"); } else { ИныеМер.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6313", "Kod", "Period2"))) { ИныеМер.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6313", "Kod", "Period2"); } else { ИныеМер.СумПред = "0"; }

                        //приравниваем объекты
                        //ВПокОПТип[] items11 = { ВПокЦелМер };
                        //РасхЦелМер.ВПокЦелМер = items11;

                        РасхЦелМер.ИныеМер = ИныеМер;
                        РасхЦелМер.СоцПом = СоцПом;
                        РасхЦелМер.ПровСемин = ПровСемин;
                        РасхЦелМер.ИныеМер = ИныеМер;

                        #endregion
                        Использовано.РасхЦелМер = РасхЦелМер;

                        #region РасхСодАУ

                        var РасхСодАУ = new ФайлДокументЦелИспИспользованоРасхСодАУ();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6320", "Kod", "Period1"))) { РасхСодАУ.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6320", "Kod", "Period1"); } else { РасхСодАУ.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6320", "Kod", "Period2"))) { РасхСодАУ.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6320", "Kod", "Period2"); } else { РасхСодАУ.СумПред = "0"; }

                        var ОплТруд = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6321", "Kod", "Period1"))) { ОплТруд.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6321", "Kod", "Period1"); } else { ОплТруд.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6321", "Kod", "Period2"))) { ОплТруд.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6321", "Kod", "Period2"); } else { ОплТруд.СумПред = "0"; }

                        var ВыплНеОТ = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6322", "Kod", "Period1"))) { ВыплНеОТ.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6322", "Kod", "Period1"); } else { ВыплНеОТ.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6322", "Kod", "Period2"))) { ВыплНеОТ.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6322", "Kod", "Period2"); } else { ВыплНеОТ.СумПред = "0"; }

                        var СлужКом = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6323", "Kod", "Period1"))) { СлужКом.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6323", "Kod", "Period1"); } else { СлужКом.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6323", "Kod", "Period2"))) { СлужКом.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6323", "Kod", "Period2"); } else { СлужКом.СумПред = "0"; }

                        var ЗданТрансп = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6324", "Kod", "Period1"))) { ЗданТрансп.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6324", "Kod", "Period1"); } else { ЗданТрансп.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6324", "Kod", "Period2"))) { ЗданТрансп.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6324", "Kod", "Period2"); } else { ЗданТрансп.СумПред = "0"; }

                        var РемОснСр = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6325", "Kod", "Period1"))) { РемОснСр.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6325", "Kod", "Period1"); } else { РемОснСр.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6325", "Kod", "Period2"))) { РемОснСр.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6325", "Kod", "Period2"); } else { РемОснСр.СумПред = "0"; }


                        var _Прочие = new ОП_ДТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6326", "Kod", "Period1"))) { _Прочие.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6326", "Kod", "Period1"); } else { _Прочие.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6326", "Kod", "Period2"))) { _Прочие.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6326", "Kod", "Period2"); } else { _Прочие.СумПред = "0"; }


                        /* var ВПокСодАУ = new ВПокОПТип();
                         ВПокСодАУ.СумОтч = " ";
                         ВПокСодАУ.НаимПок = " ";
                         ВПокСодАУ.СумПред = " ";*/

                        РасхСодАУ.ОплТруд = ОплТруд;
                        РасхСодАУ.ВыплНеОТ = ВыплНеОТ;
                        РасхСодАУ.СлужКом = СлужКом;
                        РасхСодАУ.ЗданТрансп = ЗданТрансп;
                        РасхСодАУ.РемОснСр = РемОснСр;
                        РасхСодАУ.Прочие = _Прочие;
                        //ВПокОПТип[] items_ВПокОПТип = { ВПокСодАУ };
                        //РасхСодАУ.ВПокСодАУ = items_ВПокОПТип;

                        Использовано.РасхСодАУ = РасхСодАУ;

                        #endregion


                        #region _Прочие
                        var Прочие = new ОП_ДТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6350", "Kod", "Period1"))) { Прочие.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6350", "Kod", "Period1"); } else { Прочие.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6350", "Kod", "Period2"))) { Прочие.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6350", "Kod", "Period2"); } else { Прочие.СумПред = "0"; }

                        Использовано.Прочие = Прочие;
                        //var _ВПокОП = new ВПокОПТип();
                        //_ВПокОП.СумОтч = " ";
                        //_ВПокОП.НаимПок = " ";
                        //_ВПокОП.СумПред = " ";

                        //ВПокОПТип[] items12 = { ВПокСодАУ };
                        //РасхСодАУ.ВПокСодАУ = items12;

                        //ВПокОПТип[] items13 = { ВПокОП };
                        //_Прочие.ВПокОП = items13;

                        #endregion
                      //  Использовано.Прочие = _Прочие;

                        var ПриобОСИн = new ОПТип();
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6330", "Kod", "Period1"))) { ПриобОСИн.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6330", "Kod", "Period1"); } else { ПриобОСИн.СумОтч = "0"; }
                        if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6330", "Kod", "Period2"))) { ПриобОСИн.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6330", "Kod", "Period2"); } else { ПриобОСИн.СумПред = "0"; }

                        //var ВПокИспольз = new ВПокОПТип();
                        //ВПокИспольз.СумОтч = " ";
                        //ВПокИспольз.НаимПок = " ";
                        //ВПокИспольз.СумПред = " ";
                        ЦелИсп.Использовано = Использовано;
                    }
                    #endregion

                    var ОстатНачОтч = new ОПТип();
                    if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6100", "Kod", "Period1"))) { ОстатНачОтч.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6100", "Kod", "Period1"); } else { ОстатНачОтч.СумОтч = "0"; }
                    if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6100", "Kod", "Period2"))) { ОстатНачОтч.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6100", "Kod", "Period2"); } else { ОстатНачОтч.СумПред = "0"; }
                    ЦелИсп.ОстатНачОтч = ОстатНачОтч;

                    var ОстатКонОтч = new ОПТип();
                    if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6400", "Kod", "Period1"))) { ОстатКонОтч.СумОтч = GetFields.GetCellWhere(Document, Processing, "Table", "6400", "Kod", "Period1"); } else { ОстатКонОтч.СумОтч = "0"; }
                    if (!string.IsNullOrEmpty(GetFields.GetCellWhere(Document, Processing, "Table", "6400", "Kod", "Period2"))) { ОстатКонОтч.СумПред = GetFields.GetCellWhere(Document, Processing, "Table", "6400", "Kod", "Period2"); } else { ОстатКонОтч.СумПред = "0"; }
                    ЦелИсп.ОстатКонОтч = ОстатКонОтч;


                    ФайлДокументЦелИсп[] _ЦелИсп = { ЦелИсп };
                    Документ.ЦелИсп = _ЦелИсп;
                }


                #endregion Документ_целИсп

                #region Документ_ОСВПоСчетам
                if (Document.DefinitionName == "3_100_2_Оборотно-сальдовая ведомость по счету")
                {

                    Processing.ReportMessage("3_100_2_Оборотно-сальдовая ведомость по счету");
                    var ОСВПоСчетам = new ФайлДокументОСВПоСчетам();
                    ОСВПоСчетам.Период = PeriodMonthType(GetFields.GetField(Document, Processing, "Date"));
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "PeriodYear"))) { ОСВПоСчетам.ОтчетГод = (GetFields.GetField(Document, Processing, "PeriodYear")); } else { ОСВПоСчетам.ОтчетГод = DateTime.Now.Date.Year.ToString(); }
                    var ОСВПоСчетамОКЕИ = new ФайлДокументОСВПоСчетамОКЕИ();
                    ОСВПоСчетам.ОКЕИ = ФайлДокументОСВПоСчетамОКЕИ.Item383;

                    #region ОСВПоСчету
                    Processing.ReportMessage("ОСВ по счету");
                    var ОСВПоСчетамОСВПоСчету = new ФайлДокументОСВПоСчетамОСВПоСчету();
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Num"))) { ОСВПоСчетамОСВПоСчету.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { ОСВПоСчетамОСВПоСчету.КодСчета = "-"; }
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Detail"))) { ОСВПоСчетамОСВПоСчету.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчету.НаименованиеСчета = "-"; }
                    #region СтрокаОСВ
                    Processing.ReportMessage("Строка ОСВ");
                    //List<string> notbankwords = new List<string>() { "Штрафы","Налоги и", "Поступления от", "прочие перечисления", "Перевод денежных", "аккредитив", "списание валюты", "банковские доходы", "расчеты с ", "прочие поступления", "прочие расходы", "Услуги банка", "Перевод денежных", "списание валюты", "аренда", "поступление денежных", "Денежные средства в пути", "выручка от", "е услуги", "предпродажная подготовка", "перевозки", "сертификация", "госпошлина", "дивиденды", "курсовая разница", "оплачено за", "расчеты с", "eur", "usd", "Валюта", "Валюта EUR", "Валюта USD", "Валютная сумма", "Внутренние выплаты", "Возврат", "Возврат краткосрочных кредитов", "инкассация", "Итого", "Налог на имущество", "Оплата", "Оплата Услуг", "Оплата штрафов, пени", "Получение займа", "Поступление", "Поступление платежей" };
                    //List<int> ParagraphRows = GetFields.GetTableParagraphWhereCellsNotMachText(Document, Processing, "Table", "Subkonto", notbankwords); //получаю список строк заголовков параграфов таблицы
                    #region Цикл обрабатывающий строкиОСВ


                    // ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ[] _ОСВПоСчетамОСВПоСчетуСтрокаОСВ = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ[ParagraphRows.Count];
                    ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ[] _ОСВПоСчетамОСВПоСчетуСтрокаОСВ = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ[Document.Sections[0].Field("Table").Rows.Count];
                    // int pars = ParagraphRows[0];
                    string Bank = "";
                    for (int par = 0; par < Document.Sections[0].Field("Table").Rows.Count; par++)
                    {
                        Processing.ReportMessage("Обрабатываем строку №" + par);
                        //log("GetFields.GetCellByIndex(Document, Processing, Table, Schet, par) = "+ GetFields.GetCellByIndex(Document, Processing, "Table", "Schet", par));
                        if (GetFields.GetCellByIndex(Document, Processing, "Table", "Schet", par)=="True")//(hasitem(par, ParagraphRows)) //если это новый счет
                        {
                            Bank = GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", par);
                           // log("Bank = " + Bank);
                            Processing.ReportMessage("Строка №" + par + " является параграфом");
                            
                            // int parend = par + 1;
                            // if (parend == ParagraphRows.Count) { parend = ParagraphRows.Count - 1; } //если регион последний
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВ = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ();
                            #region Счет
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВСчет();
                            if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Num"))) { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Detail"))) { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.НаименованиеСчета = "-"; }
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетТипСчета = new СтрокаОСВСТочностью3ТипСчетТипСчета();
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.ТипСчета = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетТипСчета;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "BDebet", par/*ParagraphRows[ParagraphRows.IndexOf(par)]*/).Replace(".", k()).Replace(",", k()), out decimal a);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СНД = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "BKredit",par /*ParagraphRows[ParagraphRows.IndexOf(par)*/).Replace(".", k()).Replace(",", k()), out decimal b);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СНК = b;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FlowDebet",par /*ParagraphRows[ParagraphRows.IndexOf(par)*/).Replace(".", k()).Replace(",", k()), out decimal c);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.ДО = c;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FlowKredit",par /*ParagraphRows[ParagraphRows.IndexOf(par)*/).Replace(".", k()).Replace(",", k()), out decimal d);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КО = d;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FDebet",par /*ParagraphRows[ParagraphRows.IndexOf(par)*/).Replace(".", k()).Replace(",", k()), out decimal e);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СКД = e;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FKredit",par /*ParagraphRows[ParagraphRows.IndexOf(par)*/).Replace(".", k()).Replace(",", k()), out decimal f);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СКК = f;

                            #region Субконто1
                            //СубконтоПоПредставлениюТип[] _ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1 = new СубконтоПоПредставлениюТип[ParagraphRows[ParagraphRows.IndexOf(par)]];
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1 = new СубконтоПоПредставлениюТип();

                            //Processing.ReportMessage("Итерация №" + par + "; Номер строки:" + ParagraphRows[ParagraphRows.IndexOf(par)]);
                            //Processing.ReportMessage("Субконто1");
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.Представление = Bank;//GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto",par /*ParagraphRows[ParagraphRows.IndexOf(par)*/);
                            //Processing.ReportMessage("Представление получено");
                           // log(ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.Представление);
                            if (GetFields.GetField(Document, Processing, "Detail") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.ТипСубконто = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.ТипСубконто = "-"; }
                            // Processing.ReportMessage("ТипСубконто получено");
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто1 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1;
                            Processing.ReportMessage("Тег Субконто1 закрыт");
                            #endregion Субконто1
                            //#region Субконто2
                            //var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2 = new СубконтоПоПредставлениюТип();

                            //Processing.ReportMessage("Субконто2");
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2.Представление = "-";
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2.ТипСубконто = "-";
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто2 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2;
                            //#endregion Субконто2

                            //#region Субконто3
                            //Processing.ReportMessage("Субконто3");
                            //var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3 = new СубконтоПоПредставлениюТип();
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3.Представление = "-";
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3.ТипСубконто = "-";
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто3 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3;
                            //#endregion Субконто3
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВ.Счет = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет;
                            _ОСВПоСчетамОСВПоСчетуСтрокаОСВ[par] = ОСВПоСчетамОСВПоСчетуСтрокаОСВ;

                            //pars = ParagraphRows.IndexOf(par);
                           // Processing.ReportMessage("pars=" + pars);
                        }
                        else // если это операция по счету
                        {
                            Processing.ReportMessage("Строка " + par + " является операцией по счету");
                            // int parend = par + 1;
                            // if (parend == ParagraphRows.Count) { parend = ParagraphRows.Count - 1; } //если регион последний
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВ = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ();
                            #region Счет
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВСчет();
                            if (GetFields.GetField(Document, Processing, "Num") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = "-"; }
                            if (GetFields.GetField(Document, Processing, "Detail") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = "-"; }
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетТипСчета = new СтрокаОСВСТочностью3ТипСчетТипСчета();
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.ТипСчета = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетТипСчета;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "BDebet", par).Replace(".", k()).Replace(",", k()), out decimal a);//ParagraphRows[pars]
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СНД = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "BKredit", par).Replace(".", k()).Replace(",", k()), out decimal b);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СНК = b;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FlowDebet", par).Replace(".", k()).Replace(",", k()), out decimal c);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.ДО = c;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FlowKredit", par).Replace(".", k()).Replace(",", k()), out decimal d);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КО = d;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FDebet", par).Replace(".", k()).Replace(",", k()), out decimal e);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СКД = e;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FKredit", par).Replace(".", k()).Replace(",", k()), out decimal f);
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СКК = f;

                            #region Субконто1
                            //СубконтоПоПредставлениюТип[] _ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1 = new СубконтоПоПредставлениюТип[ParagraphRows[par]];
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1 = new СубконтоПоПредставлениюТип();

                            // Processing.ReportMessage("Итерация №" + par + "; Номер строки:" + ParagraphRows[pars]);
                            Processing.ReportMessage("Субконто1");
                            // Processing.ReportMessage("[pars] = " + pars);
                            //Processing.ReportMessage("ParagraphRows[pars] = "+ ParagraphRows[pars]);

                            if (!String.IsNullOrEmpty(Bank/*GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", par ParagraphRows[pars]*/)) { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.Представление = Bank/*GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", parParagraphRows[pars]*/; } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.Представление = "-"; }
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.Представление = GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", ParagraphRows[pars]); 
                            // Processing.ReportMessage("Представление получено");
                            if (GetFields.GetField(Document, Processing, "Detail") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.ТипСубконто = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.ТипСубконто = "-"; }
                            // Processing.ReportMessage("ТипСубконто получено");
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто1 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1;
                            Processing.ReportMessage("Тег Субконто1 закрыт");
                            #endregion Субконто1
                            #region Субконто2
                            var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2 = new СубконтоПоПредставлениюТип();

                            Processing.ReportMessage("Субконто2");

                            if (!String.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", par)))
                            {
                                ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2.Представление = GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", par); ;
                                ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2.ТипСубконто = "-";
                                ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто2 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2;
                            }
                            #endregion Субконто2
                            //#region Субконто3
                            //Processing.ReportMessage("Субконто3");
                            //var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3 = new СубконтоПоПредставлениюТип();
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3.Представление = "-";
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3.ТипСубконто = "-";
                            //ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто3 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3;
                            //#endregion Субконто3

                            #endregion Счет
                            ОСВПоСчетамОСВПоСчетуСтрокаОСВ.Счет = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет;
                            _ОСВПоСчетамОСВПоСчетуСтрокаОСВ[par] = ОСВПоСчетамОСВПоСчетуСтрокаОСВ;

                        }
                        #endregion
                    }



                    //for (int par = 0; par < ParagraphRows.Count; par++)
                    //{
                    //    Processing.ReportMessage("Идем по параграфам. Обрабатываем параграф в строке №" + ParagraphRows[par]);
                    //    int parend = par + 1;
                    //    if (parend == ParagraphRows.Count) { parend = ParagraphRows.Count - 1; } //если регион последний
                    //    var ОСВПоСчетамОСВПоСчетуСтрокаОСВ = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВ();
                    //    #region Счет
                    //    var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет = new ФайлДокументОСВПоСчетамОСВПоСчетуСтрокаОСВСчет();
                    //    if (GetFields.GetField(Document, Processing, "Num") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = "-"; }
                    //    if (GetFields.GetField(Document, Processing, "Detail") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КодСчета = "-"; }
                    //    var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетТипСчета = new СтрокаОСВСТочностью3ТипСчетТипСчета();
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.ТипСчета = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетТипСчета;
                    //    Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "BDebet", ParagraphRows[par]).Replace(".", ","), out decimal a);
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СНД = a;
                    //    Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "BKredit", ParagraphRows[par]).Replace(".", ","), out decimal b);
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СНК = b;
                    //    Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FlowDebet", ParagraphRows[par]).Replace(".", ","), out decimal c);
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.ДО = c;
                    //    Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FlowKredit", ParagraphRows[par]).Replace(".", ","), out decimal d);
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.КО = d;
                    //    Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FDebet", ParagraphRows[par]).Replace(".", ","), out decimal e);
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СКД = e;
                    //    Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "FKredit", ParagraphRows[par]).Replace(".", ","), out decimal f);
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.СКК = f;

                    //    #region Субконто1
                    //    //СубконтоПоПредставлениюТип[] _ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1 = new СубконтоПоПредставлениюТип[ParagraphRows[par]];
                    //    var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1 = new СубконтоПоПредставлениюТип();

                    //    Processing.ReportMessage("Итерация №" + par + "; Номер строки:" + ParagraphRows[par]);
                    //    Processing.ReportMessage("Субконто1");
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.Представление = GetFields.GetCellByIndex(Document, Processing, "Table", "Subkonto", ParagraphRows[par]); ;
                    //    Processing.ReportMessage("Представление получено");
                    //    if (GetFields.GetField(Document, Processing, "Detail") != "") { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.ТипСубконто = (GetFields.GetField(Document, Processing, "Detail")); } else { ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1.ТипСубконто = "-"; }
                    //    Processing.ReportMessage("ТипСубконто получено");
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто1 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто1;
                    //    Processing.ReportMessage("Тег Субконто1 закрыт");
                    //    #endregion Субконто1
                    //    #region Субконто2
                    //    var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2 = new СубконтоПоПредставлениюТип();

                    //    Processing.ReportMessage("Субконто2");
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2.Представление = "-";
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2.ТипСубконто = "-";
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто2 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто2;
                    //    #endregion Субконто2
                    //    #region Субконто3
                    //    Processing.ReportMessage("Субконто3");
                    //    var ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3 = new СубконтоПоПредставлениюТип();
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3.Представление = "-";
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3.ТипСубконто = "-";
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет.Субконто3 = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчетСубконто3;
                    //    #endregion Субконто3
                    //    ОСВПоСчетамОСВПоСчетуСтрокаОСВ.Счет = ОСВПоСчетамОСВПоСчетуСтрокаОСВСчет;
                    //    _ОСВПоСчетамОСВПоСчетуСтрокаОСВ[par] = ОСВПоСчетамОСВПоСчетуСтрокаОСВ;
                    //    #endregion Счет
                    //}


                    ОСВПоСчетамОСВПоСчету.СтрокаОСВ = _ОСВПоСчетамОСВПоСчетуСтрокаОСВ;
                    #endregion Цикл обрабатывающий строкиОСВ
                    #endregion СтрокаОСВ
                    ФайлДокументОСВПоСчетамОСВПоСчету[] _ОСВПоСчетамОСВПоСчету = { ОСВПоСчетамОСВПоСчету };
                    ОСВПоСчетам.ОСВПоСчету = _ОСВПоСчетамОСВПоСчету;
                    #endregion ОСВПоСчету
                    ФайлДокументОСВПоСчетам[] _ОСВПоСчетам = { ОСВПоСчетам };
                    Документ.ОСВПоСчетам = _ОСВПоСчетам;
                }

                #endregion

                #region Документ_Денежные средства

                if (Document.DefinitionName == "Денежные средства")
                {
                    var ДенежныеСредства = new ФайлДокументДенежныеСредства();
                    ДенежныеСредства.Период = PeriodMonthType(GetFields.GetField(Document, Processing, "Date"));
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "PeriodYear"))) { ДенежныеСредства.ОтчетГод = (GetFields.GetField(Document, Processing, "PeriodYear")); } else { ДенежныеСредства.ОтчетГод = DateTime.Now.Date.Year.ToString(); }
                    var ДенежныеСредстваОКЕИ = new ФайлДокументДенежныеСредстваОКЕИ();
                    ДенежныеСредства.ОКЕИ =ФайлДокументДенежныеСредстваОКЕИ.Item384;
                    Processing.ReportMessage("Начинаю обрабатывать счета");

                    var ДенежныеСредстваАнализ_51 = new АнализСчетаСТочностью3ТипСтрокаАнализа();

                    List<int> ParagraphRows = GetFields.GetTableParagraph(Document, Processing, "Table", "Acc");
                    Processing.ReportMessage("Получил список счетов");
                    for (int par = 0; par < ParagraphRows.Count; par++)
                    {
                        Processing.ReportMessage("Обрабатываю счет №" + par + " в строке №" + ParagraphRows[par]);
                        var ДенежныеСредстваАнализ_51Счет = new АнализСчетаСТочностью3ТипСтрокаАнализаСчет();
                        //ДенежныеСредстваАнализ_51Счет.КодСчета = "51";
                        if (GetFields.GetField(Document, Processing, "Num") != "") { ДенежныеСредстваАнализ_51Счет.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { ДенежныеСредстваАнализ_51Счет.КодСчета = "-"; }
                        //ДенежныеСредстваАнализ_51Счет.НаименованиеСчета = "Активный";
                        if (GetFields.GetField(Document, Processing, "Detail") != "") { ДенежныеСредстваАнализ_51Счет.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { ДенежныеСредстваАнализ_51Счет.КодСчета = "-"; }
                        ДенежныеСредстваАнализ_51Счет.ТипСчета = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                        //ДенежныеСредстваАнализ_51Счет.КодСчетаР = "000000";
                        //ДенежныеСредстваАнализ_51Счет.ТипСчетаР = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                        // ДенежныеСредстваАнализ_51Счет.КодСчетаРР = "000000";
                        //ДенежныеСредстваАнализ_51Счет.ТипСчетаРР = СтрокаОСВСТочностью3ТипСчетТипСчета.А;


                        int parend = par + 1;
                        if (parend >= ParagraphRows.Count)
                        {
                            parend = ParagraphRows.Count - 1;
                        }
                        Processing.ReportMessage("ParagraphRows.Count =" + ParagraphRows.Count + " par =" + par + " parend=" + parend);
                        Processing.ReportMessage("ParagraphRows[par]=" + ParagraphRows[par]);
                        Processing.ReportMessage(" ParagraphRows[parend]=" + ParagraphRows[parend]);

                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "нач.сальдо", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal a);
                        ДенежныеСредстваАнализ_51Счет.СНД = a;
                        Processing.ReportMessage("Текст СНД = " + GetFields.GetCellByIndexWhere(Document, Processing, "Table", "нач.сальдо", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]));
                        Processing.ReportMessage("ДенежныеСредстваАнализ_51Счет.СНД = " + a);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "нач.сальдо", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal b);
                        ДенежныеСредстваАнализ_51Счет.СНК = b;
                        Processing.ReportMessage("Текст СНК = " + GetFields.GetCellByIndexWhere(Document, Processing, "Table", "нач.сальдо", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]));
                        Processing.ReportMessage("ДенежныеСредстваАнализ_51Счет.СНК = " + b);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "оборот", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal c);
                        ДенежныеСредстваАнализ_51Счет.ДО = c;
                        Processing.ReportMessage("ДенежныеСредстваАнализ_51Счет.ДО = " + c);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "оборот", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal d);
                        ДенежныеСредстваАнализ_51Счет.КО = d;
                        Processing.ReportMessage("ДенежныеСредстваАнализ_51Счет.КО  = " + d);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "кон.сальдо", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal e);
                        ДенежныеСредстваАнализ_51Счет.СКД = e;
                        Processing.ReportMessage("ДенежныеСредстваАнализ_51Счет.КД  = " + e);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "кон.сальдо", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal f);
                        ДенежныеСредстваАнализ_51Счет.СКК = f;
                        Processing.ReportMessage("ДенежныеСредстваАнализ_51Счет.СКК  = " + f);
                            // ДенежныеСредстваАнализ_51Счет.ПериодГод = DateTime.Now.Date.Year.ToString();
                        if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Date")) && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out var dateValue2)) { ДенежныеСредстваАнализ_51Счет.ПериодГод = dateValue2.Year.ToString(); } else { ДенежныеСредстваАнализ_51Счет.ПериодГод = DateTime.Now.Year.ToString(); }

                        ДенежныеСредстваАнализ_51Счет.ПериодМесяц = DateTime.Now.Date.Month.ToString();

                        var ДенежныеСредстваСубконто = new АнализСчетаСТочностью3ТипСтрокаАнализаСчетСубконто();
                        var ДенежныеСредстваСубконтоItem = new БанковскийСчетТип();
                        ДенежныеСредстваСубконтоItem.Наименование = GetFields.GetCellByIndex(Document, Processing, "Table", "Acc", ParagraphRows[par]);
                        //ДенежныеСредстваСубконтоItem.НаименованиеБанка = "";
                        //ДенежныеСредстваСубконтоItem.БИК = "";
                        // ДенежныеСредстваСубконтоItem.НомерСчета = "";

                        АнализСчетаСТочностью3ТипСтрокаАнализаСчетСубконто[] _ДенежныеСредстваСубконто = { ДенежныеСредстваСубконто };
                        ДенежныеСредстваАнализ_51Счет.Субконто = _ДенежныеСредстваСубконто;

                        АнализСчетаСТочностью3ТипСтрокаАнализаСчет _ДенежныеСредстваАнализ_51Счет = ДенежныеСредстваАнализ_51Счет;
                        ДенежныеСредстваАнализ_51.Счет = _ДенежныеСредстваАнализ_51Счет;

                        Processing.ReportMessage("получаю список корреспонденции");
                        Regex rgx = new Regex(@"\d{2,3}");
                        List<int> accslist = GetFields.GetTableRowsWhereRegex(Document, Processing, "Table", "CorrAcc", rgx, par, parend);
                        Processing.ReportMessage("accslist.Count =" + accslist.Count);
                        for (int i = 0; i < accslist.Count; i++)
                        {
                            Processing.ReportMessage("получаю корреспонденцию из строки №" + accslist[i]);
                            var ДенежныеСредстваКорреспонденция = new АнализСчетаСТочностью3ТипСтрокаАнализаКорреспонденция();
                            ДенежныеСредстваКорреспонденция.КодСчета = GetFields.GetCellByIndex(Document, Processing, "Table", "CorrAcc", accslist[i]);
                            ДенежныеСредстваКорреспонденция.НаименованиеСчета = "-";
                            ДенежныеСредстваКорреспонденция.ТипСчета = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                            //ДенежныеСредстваКорреспонденция.КодСчетаР = "-";
                            // ДенежныеСредстваКорреспонденция.ТипСчетаР = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                            //ДенежныеСредстваКорреспонденция.КодСчетаРР = "-";
                            //ДенежныеСредстваКорреспонденция.ТипСчетаРР = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "DebAcc", accslist[i]).Replace(".", k()).Replace(",", k()), out decimal z);
                            ДенежныеСредстваКорреспонденция.ДО = z;
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "KredAcc", accslist[i]).Replace(".", k()).Replace(",", k()), out decimal x);
                            ДенежныеСредстваКорреспонденция.КО = x;

                            АнализСчетаСТочностью3ТипСтрокаАнализаКорреспонденция[] _ДенежныеСредстваКорреспонденция = { ДенежныеСредстваКорреспонденция };
                            ДенежныеСредстваАнализ_51.Корреспонденция = _ДенежныеСредстваКорреспонденция;
                        }

                    }
                    АнализСчетаСТочностью3ТипСтрокаАнализа[] _ДенежныеСредстваАнализ_51 = { ДенежныеСредстваАнализ_51 };
                    ДенежныеСредства.Анализ_51 = _ДенежныеСредстваАнализ_51;

                    ФайлДокументДенежныеСредства[] _ДенежныеСредства = { ДенежныеСредства };
                    Документ.ДенежныеСредства = _ДенежныеСредства;
                }
                #endregion Документ_Денежные средства

                #region Документ_АнализСчетов
                if (Document.DefinitionName == "3_55_Расшифровка сумм по аналитическим счетам")
                {
                    Processing.ReportMessage("Начинаю формирование ХМЛ по документу - Анализ счетов");
                    var АнализСчетов = new ФайлДокументАнализСчетов();
                    АнализСчетов.Период = PeriodMonthType(GetFields.GetField(Document, Processing, "Date"));
                    if (GetFields.GetField(Document, Processing, "PeriodYear") != "") { АнализСчетов.ОтчетГод = (GetFields.GetField(Document, Processing, "PeriodYear")); } else { АнализСчетов.ОтчетГод = DateTime.Now.Date.Year.ToString(); }
                    //АнализСчетов.ОтчетГод = "2019";
                    var АнализСчетовОКЕИ = new ФайлДокументАнализСчетовОКЕИ();
                    АнализСчетов.ОКЕИ = ФайлДокументАнализСчетовОКЕИ.Item383;

                    #region Анализ счета
                    Processing.ReportMessage("тег: Анализ счета");
                    var АнализСчетовАнализСчета = new ФайлДокументАнализСчетовАнализСчета();
                    if (GetFields.GetField(Document, Processing, "Num") != "") { АнализСчетовАнализСчета.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { АнализСчетовАнализСчета.КодСчета = "-"; }
                    if (GetFields.GetField(Document, Processing, "Detail") != "") { АнализСчетовАнализСчета.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { АнализСчетовАнализСчета.НаименованиеСчета = "-"; }


                    #region СтрокаАнализа 
                    Processing.ReportMessage("тег: Строка анализа");
                    //цикл
                    List<int> ParagraphRows = GetFields.GetTableParagraphByCheckMark(Document, Processing, "Table", "Schet"); //получаю список строк заголовков параграфов таблицы
                    ФайлДокументАнализСчетовАнализСчетаСтрокаАнализа[] _АнализСчетаСтрокаАнализа = new ФайлДокументАнализСчетовАнализСчетаСтрокаАнализа[ParagraphRows.Count];
                    for (int par = 0; par < ParagraphRows.Count; par++)
                    {
                        Processing.ReportMessage("Идем по параграфам. Обрабатываем параграф №" + par);

                        int parend = par + 1;
                        if (parend == ParagraphRows.Count) { parend = ParagraphRows.Count - 1; } //если регион последний
                        #region счет
                        Processing.ReportMessage("тег: счет");
                        var АнализСчетаСтрокаАнализа = new ФайлДокументАнализСчетовАнализСчетаСтрокаАнализа();
                        var АнализСчетаСтрокаАнализаСчет = new ФайлДокументАнализСчетовАнализСчетаСтрокаАнализаСчет();
                        
                        if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Num"))) { АнализСчетаСтрокаАнализаСчет.КодСчета = (GetFields.GetField(Document, Processing, "Num")); } else { АнализСчетаСтрокаАнализаСчет.КодСчета = "-"; }
                        //АнализСчетаСтрокаАнализаСчет.КодСчета = "";
                        if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Detail"))) { АнализСчетаСтрокаАнализаСчет.НаименованиеСчета = (GetFields.GetField(Document, Processing, "Detail")); } else { АнализСчетаСтрокаАнализаСчет.НаименованиеСчета = "-"; }
                        //АнализСчетаСтрокаАнализаСчет.НаименованиеСчета = "";
                        АнализСчетаСтрокаАнализаСчет.ТипСчета = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "нач", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal a);
                        Processing.ReportMessage("a=" + a);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "нач", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal b);
                        Processing.ReportMessage("b=" + b);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "оборот", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal c);
                        Processing.ReportMessage("c=" + c);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "оборот", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal d);
                        Processing.ReportMessage("d=" + d);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "кон", "CorrAcc", "DebAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal e);
                        Processing.ReportMessage("e=" + e);
                        Decimal.TryParse(GetFields.GetCellByIndexWhere(Document, Processing, "Table", "кон", "CorrAcc", "KredAcc", ParagraphRows[par], ParagraphRows[parend]).Replace(".", k()).Replace(",", k()), out decimal f);
                        Processing.ReportMessage("f=" + f);
                        АнализСчетаСтрокаАнализаСчет.СНД = a;
                        АнализСчетаСтрокаАнализаСчет.СНК = b;
                        АнализСчетаСтрокаАнализаСчет.ДО = c;
                        АнализСчетаСтрокаАнализаСчет.КО = d;
                        АнализСчетаСтрокаАнализаСчет.СКД = e;
                        АнализСчетаСтрокаАнализаСчет.СКК = f;

                        #region субконто1
                        Processing.ReportMessage("субконто1");
                        if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "Acc", ParagraphRows[par])))
                        {
                            // var АнализСчетаСтрокаАнализаСчетСубконто1 = new АнализСчетаСТочностью3ТипСтрокаАнализаСчетСубконто();
                            var АнализСчетаСчетСубконтоПредставление1 = new СубконтоПоПредставлениюТип();
                            var АнализСчетаСтрокаАнализаСчетСубконто1 = new АнализСчетаСТочностью3ТипСтрокаАнализаСчетСубконто();
                            АнализСчетаСчетСубконтоПредставление1.Представление = GetFields.GetCellByIndex(Document, Processing, "Table", "Acc", ParagraphRows[par]); ;
                            АнализСчетаСчетСубконтоПредставление1.ТипСубконто = "-";
                            АнализСчетаСтрокаАнализаСчет.Субконто1 = АнализСчетаСчетСубконтоПредставление1;
                        }

                        #endregion субконто1

                        //#region субконто2
                        //Processing.ReportMessage("субконто2");
                        //var АнализСчетаСтрокаАнализаСчетСубконто2 = new АнализСчетаСТочностью3ТипСтрокаАнализаСчетСубконто();
                        //var АнализСчетаСчетСубконтоПредставление2 = new СубконтоПоПредставлениюТип();
                        //АнализСчетаСчетСубконтоПредставление2.Представление = "-";
                        //АнализСчетаСчетСубконтоПредставление2.ТипСубконто = "-";
                        //АнализСчетаСтрокаАнализаСчет.Субконто2 = АнализСчетаСчетСубконтоПредставление2;
                        //#endregion субконто2

                        //#region субконто3
                        //Processing.ReportMessage("субконто3");
                        //var АнализСчетаСтрокаАнализаСчетСубконто3 = new АнализСчетаСТочностью3ТипСтрокаАнализаСчетСубконто();
                        //var АнализСчетаСчетСубконтоПредставление3 = new СубконтоПоПредставлениюТип();
                        //АнализСчетаСчетСубконтоПредставление3.Представление = "-";
                        //АнализСчетаСчетСубконтоПредставление3.ТипСубконто = "-";
                        //АнализСчетаСтрокаАнализаСчет.Субконто3 = АнализСчетаСчетСубконтоПредставление3;
                        //#endregion субконто3

                        АнализСчетаСтрокаАнализа.Счет = АнализСчетаСтрокаАнализаСчет;

                        #endregion счет

                        #region корреспонденция
                        Processing.ReportMessage("тег: Корреспонденция");

                        Regex rgx = new Regex(@"\d{2,3}");
                        List<int> accslist = GetFields.GetTableRowsWhereRegex(Document, Processing, "Table", "CorrAcc", rgx, ParagraphRows[par], ParagraphRows[parend]);

                        ФайлДокументАнализСчетовАнализСчетаСтрокаАнализаКорреспонденция[] _АнализСчетаСтрокаАнализаКорреспонденция = new ФайлДокументАнализСчетовАнализСчетаСтрокаАнализаКорреспонденция[accslist.Count];
                        Processing.ReportMessage("Количество счетов  корреспонденции =" + accslist.Count);
                        for (int i = 0; i < accslist.Count; i++)
                        {
                            var АнализСчетаСтрокаАнализаКорреспонденция = new ФайлДокументАнализСчетовАнализСчетаСтрокаАнализаКорреспонденция();
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "CorrAcc", accslist[i]))) { АнализСчетаСтрокаАнализаКорреспонденция.КодСчета = GetFields.GetCellByIndex(Document, Processing, "Table", "CorrAcc", accslist[i]); } else { АнализСчетаСтрокаАнализаКорреспонденция.КодСчета = "-"; }
                            АнализСчетаСтрокаАнализаКорреспонденция.НаименованиеСчета = "-";
                            АнализСчетаСтрокаАнализаКорреспонденция.ТипСчета = СтрокаОСВСТочностью3ТипСчетТипСчета.А;
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "DebAcc", accslist[i]).Replace(".", k()).Replace(",", k()), out decimal z);
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "KredAcc", accslist[i]).Replace(".", k()).Replace(",", k()), out decimal x);
                            АнализСчетаСтрокаАнализаКорреспонденция.ДО = z;
                            АнализСчетаСтрокаАнализаКорреспонденция.КО = x;
                            Processing.ReportMessage("Закрываю тег корреспонденции");

                            Processing.ReportMessage("i = " + par + " accslist.Count= " + accslist.Count);
                            _АнализСчетаСтрокаАнализаКорреспонденция[i] = АнализСчетаСтрокаАнализаКорреспонденция;

                        }
                        АнализСчетаСтрокаАнализа.Корреспонденция = _АнализСчетаСтрокаАнализаКорреспонденция;
                        #endregion корреспонденция
                        Processing.ReportMessage("Закрываю тег СтрокаАнализа");

                        Processing.ReportMessage("par = " + par + " ParagraphRows.Count= " + ParagraphRows.Count);
                        _АнализСчетаСтрокаАнализа[par] = АнализСчетаСтрокаАнализа;
                        АнализСчетовАнализСчета.СтрокаАнализа = _АнализСчетаСтрокаАнализа;
                        #endregion СтрокаАнализа
                    }

                    ФайлДокументАнализСчетовАнализСчета[] _АнализСчетовАнализСчета = { АнализСчетовАнализСчета };
                    АнализСчетов.АнализСчета = _АнализСчетовАнализСчета;

                    #endregion Анализ счета
                    ФайлДокументАнализСчетов[] _АнализСчетов = { АнализСчетов };
                    Документ.АнализСчетов = _АнализСчетов;
                }
                #endregion Документ_АнализСчетов

                #region Документ_РасшДебКредЗад
                if (Document.DefinitionName == "3_42_1_Расшифровка кредиторской задолженности" || Document.DefinitionName == "3_42_2_Расшифровка дебиторской задолженности")
                {
                    Processing.ReportMessage("Начинаю формирование ХМЛ по документу - РасшДебКредЗад");
                    var ДебКредЗадолж = new ФайлДокументДебКредЗадолж();
                    ДебКредЗадолж.Период = PeriodQuarterType(GetFields.GetField(Document, Processing, "Date"));
                   
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Date")) && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out var dateValue2)) { ДебКредЗадолж.ОтчетГод = dateValue2.Year.ToString(); } else { ДебКредЗадолж.ОтчетГод = DateTime.Now.Year.ToString(); }

                    ДебКредЗадолж.ОКЕИ = ФайлДокументДебКредЗадолжОКЕИ.Item384;



                    #region Интервал
                    ////тут должен быть цикл 
                    //var i = 0;
                    //var Интервал = new ФайлДокументДебКредЗадолжИнтервалыИнтервал();
                    //Интервал.Код = "000000";
                    //Интервал.Начало = "-";
                    //Интервал.Окончание = "-";
                    //// ФайлДокументДебКредЗадолжИнтервалыИнтервал[][] _Интервал = { Интервал } ;
                    ////ДебКредЗадолж.Интервалы = _Интервал;


                    #endregion Интервал

                    #region ДебДсрЗадолжВид
                    //Processing.ReportMessage(" ДебДсрЗадолжВид");
                    //var ДебДсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид();
                    //ДебДсрЗадолжВид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид.Item1;

                    //#region Задолж
                    //var ЗадолжВидЗадолж = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж();

                    //#region Контрагент
                    //var ЗадолжВидЗадолжКонтрагент = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент();
                    //ЗадолжВидЗадолжКонтрагент.Наим = "qwerty";
                    //ЗадолжВидЗадолжКонтрагент.ИНН = "1214567890";
                    //ЗадолжВидЗадолжКонтрагент.КПП = "121456789";
                    //ЗадолжВидЗадолжКонтрагент.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗадолжКонтрагент.Общая = "121";

                    //#region Договор
                    //var ЗадолжВидЗфдолжДоговор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор();
                    //ЗадолжВидЗфдолжДоговор.Наим = "qwerty";
                    //ЗадолжВидЗфдолжДоговор.Номер = "qwerty";
                    //ЗадолжВидЗфдолжДоговор.Дата = "01.01.2019";
                    //ЗадолжВидЗфдолжДоговор.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗфдолжДоговор.Общая = "12145";
                    //ЗадолжВидЗфдолжДоговор.Просроч = "12145";

                    //#region ПоСроку
                    //var ЗадолжВидЗадолжКонтрагентДоговорПоСроку = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку();
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку.Код = "12145";
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку.Код = "12145";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку[] _ЗадолжВидЗадолжКонтрагентДоговорПоСроку = { ЗадолжВидЗадолжКонтрагентДоговорПоСроку };
                    //ЗадолжВидЗфдолжДоговор.ПоСроку = _ЗадолжВидЗадолжКонтрагентДоговорПоСроку;
                    //#endregion Договор

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор[] _ЗадолжВидЗфдолжДоговор = { ЗадолжВидЗфдолжДоговор };
                    //ЗадолжВидЗадолжКонтрагент.Договор = _ЗадолжВидЗфдолжДоговор;

                    //#endregion

                    //// ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент[] _ЗадолжВидЗадолжКонтрагент = { ЗадолжВидЗадолжКонтрагент };
                    //ЗадолжВидЗадолж.Контрагент = ЗадолжВидЗадолжКонтрагент;

                    //#endregion Задолж

                    //#region Итого
                    //var ЗадолжВидИтого = new ФайлДокументДебКредЗадолжЗадолжВидИтого();
                    //ЗадолжВидИтого.Общая = "245";
                    //ЗадолжВидИтого.Просроч = "245";

                    //#region ПоСроку
                    //var ЗадолжВидИтогоПоСроку = new ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку();
                    //ЗадолжВидИтогоПоСроку.Код = "11145";
                    //ЗадолжВидИтогоПоСроку.Задолж = "12145";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку[] _ЗадолжВидИтогоПоСроку = { ЗадолжВидИтогоПоСроку };
                    //ЗадолжВидИтого.ПоСроку = _ЗадолжВидИтогоПоСроку;
                    //#endregion Итого

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолж[] _ЗадолжВидЗадолж = { ЗадолжВидЗадолж };
                    //ДебДсрЗадолжВид.Задолж = _ЗадолжВидЗадолж;
                    //// ФайлДокументДебКредЗадолжЗадолжВидИтого[] _ЗадолжВидИтого = { ЗадолжВидИтого };
                    //ДебДсрЗадолжВид.Итого = ЗадолжВидИтого;
                    //ФайлДокументДебКредЗадолжЗадолжВид[] _ДебДсрЗадолжВид = { ДебДсрЗадолжВид };
                    //ДебКредЗадолж.ДебДср = _ДебДсрЗадолжВид;

                    //Processing.ReportMessage(" ДебДсрЗадолжВид - закончили");
                    #endregion Конец_ДебДсрЗадолжВид

                    #region ДебКсрЗадолжВид
                    if (Document.DefinitionName == "3_42_2_Расшифровка дебиторской задолженности")
                    {
                        Processing.ReportMessage(" ДебКсрЗадолжВид");
                        var ДебКсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид1();
                        ДебКсрЗадолжВид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид1.Item1;
                        #region Задолж
                        Processing.ReportMessage(" Задолж");
                        List<string> notprochee = new List<string>() { "Прочее", "Прочие", "Итого" };
                        List<int> AccsRows = GetFields.GetTableParagraphWhereCellsNotMachText(Document, Processing, "Table", "Name", notprochee);
                        Processing.ReportMessage("AccsRows = " + AccsRows.Count);
                        Regex rgx = new Regex(@".*(Прочие|Прочее).*", RegexOptions.IgnoreCase);
                        List<int> ProchRows = GetFields.GetTableRowsWhereRegex(Document, Processing, "Table", "Name", rgx, 0, Document.Sections[0].Field("Table").Rows.Count);
                        Processing.ReportMessage("ProchRows = " + ProchRows.Count);
                        List<int> AllRows = new List<int>() { };
                        foreach (int it in AccsRows)
                            AllRows.Add(it);
                        foreach (int it in ProchRows)
                            AllRows.Add(it);

                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж1[] ЗадолжВидЗадолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж1[AllRows.Count]; // Массив Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент1[] ЗадолжВидЗадолжКонтрагент3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент1[AccsRows.Count]; // массив Контрагентов
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолжПрочее[] ЗадолжВидПрочее = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжПрочее[ProchRows.Count]; //массив Прочее

                        Processing.ReportMessage("Контрагенты");
                        for (var l = 0; l < AccsRows.Count; l++) //цикл по Контрагентам
                        {
                            var _ЗадолжВидЗадолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж1();
                            var _ЗадолжКонтрагент3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент1();
                            #region данные
                            Processing.ReportMessage("Обрабатываем строку - " + l);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "Name", AccsRows[l]))) { _ЗадолжКонтрагент3.Наим = GetFields.GetCellByIndex(Document, Processing, "Table", "Name", AccsRows[l]).ToString(); } else { _ЗадолжКонтрагент3.Наим = "-"; };
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "INN", AccsRows[l]))) { _ЗадолжКонтрагент3.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table", "INN", AccsRows[l]).ToString(); } else { };
                            //ЗадолжВидЗадолжКонтрагент3.КПП = "123456789";
                            //ЗадолжВидЗадолжКонтрагент3.ДатаВозн = "01.01.2019";
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "Sum", AccsRows[l]), out decimal a);
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "SumOverdue", AccsRows[l]), out decimal b);
                            // _ЗадолжКонтрагент3.Общая = a.ToString();
                            #endregion данные

                            #region Договор

                            Processing.ReportMessage("Договор");

                            var ЗадолжВидЗфдолж3Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор1();
                            //ЗадолжВидЗфдолж3Договор.Наим = "qwerty";
                            //ЗадолжВидЗфдолж3Договор.Номер = "qwerty";
                            //ЗадолжВидЗфдолж3Договор.Дата = "01.01.2019";
                            //ЗадолжВидЗфдолж3Договор.ДатаВозн = "01.01.2019";
                            ЗадолжВидЗфдолж3Договор.Общая = a;//Decimal.ToInt32(a).ToString();
                            ЗадолжВидЗфдолж3Договор.Просроч = b;// Decimal.ToInt32(b).ToString();
                            //ЗадолжВидЗфдолж3Договор.Просроч = "12345";

                            #region ПоСроку
                            //var ЗадолжВидЗадолжКонтрагентДоговорПоСроку3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку1();
                            //ЗадолжВидЗадолжКонтрагентДоговорПоСроку3.Код = "12345";
                            //ЗадолжВидЗадолжКонтрагентДоговорПоСроку3.Код = "12345";
                            #endregion ПоСроку

                            //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку1[] _ЗадолжВидЗадолжКонтрагентДоговорПоСроку3 = { ЗадолжВидЗадолжКонтрагентДоговорПоСроку3 };
                            //ЗадолжВидЗфдолж3Договор.ПоСроку = _ЗадолжВидЗадолжКонтрагентДоговорПоСроку3;

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор1[] _ЗадолжВидЗфдолж3Договор = { ЗадолжВидЗфдолж3Договор };
                            _ЗадолжКонтрагент3.Договор = _ЗадолжВидЗфдолж3Договор;

                            #endregion Договор

                            ЗадолжВидЗадолжКонтрагент3[l] = _ЗадолжКонтрагент3;
                            _ЗадолжВидЗадолж3.Item = ЗадолжВидЗадолжКонтрагент3[l];
                            ЗадолжВидЗадолж3[l] = _ЗадолжВидЗадолж3;
                            ДебКсрЗадолжВид.Задолж = ЗадолжВидЗадолж3;
                        }


                        for (var h = 0; h < ProchRows.Count; h++) //цикл по прочее
                        {
                            var _ЗадолжВидЗадолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж1();
                            var _ЗадолжПрочее = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжПрочее();
                            #region данные
                            if (!String.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "Name", ProchRows[h]))) { _ЗадолжПрочее.Наим = GetFields.GetCellByIndex(Document, Processing, "Table", "Name", ProchRows[h]).ToString(); } else { _ЗадолжПрочее.Наим = "-"; };
                            //  ЗадолжВидПрочее[l].Наим = GetFields.GetCellByIndex(Document, Processing, "Table", "Name", ProchRows[l]);
                            //ЗадолжВидПрочее.ДатаВозн = "";
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "Sum", ProchRows[h]), out decimal a);
                            _ЗадолжПрочее.Общая = a;//Decimal.ToInt32(a).ToString();
                            #endregion данные

                            ЗадолжВидПрочее[h] = _ЗадолжПрочее;
                            _ЗадолжВидЗадолж3.Item = ЗадолжВидПрочее[h];
                            ЗадолжВидЗадолж3[AccsRows.Count + h] = _ЗадолжВидЗадолж3;
                            ДебКсрЗадолжВид.Задолж = ЗадолжВидЗадолж3;
                        }
                        Processing.ReportMessage("закрыли ЗадолжВидЗадолж3");
                        #endregion Задолж


                        ФайлДокументДебКредЗадолжЗадолжВид1[] _КредКсрЗадолжВид = { ДебКсрЗадолжВид };
                        ДебКредЗадолж.ДебКср = _КредКсрЗадолжВид;
                        Processing.ReportMessage("закрыли ДебКср");

                    }
                    //var ДебКсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид1();
                    //ДебКсрЗадолжВид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид1.Item1;

                    //#region Задолж
                    //var ЗадолжВидЗадолж1 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж1();

                    //#region Контрагент
                    //var ЗадолжВидЗадолжКонтрагент1 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент1();
                    //ЗадолжВидЗадолжКонтрагент1.Наим = "qwerty";
                    //ЗадолжВидЗадолжКонтрагент1.ИНН = "1214567890";
                    //ЗадолжВидЗадолжКонтрагент1.КПП = "121456789";
                    //ЗадолжВидЗадолжКонтрагент1.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗадолжКонтрагент1.Общая = "121";

                    //#region Договор
                    //var ЗадолжВидЗфдолж1Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор1();
                    //ЗадолжВидЗфдолж1Договор.Наим = "qwerty";
                    //ЗадолжВидЗфдолж1Договор.Номер = "qwerty";
                    //ЗадолжВидЗфдолж1Договор.Дата = "01.01.2019";
                    //ЗадолжВидЗфдолж1Договор.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗфдолж1Договор.Общая = "12145";
                    //ЗадолжВидЗфдолж1Договор.Просроч = "12145";

                    //#region ПоСроку
                    //var ЗадолжВидЗадолжКонтрагентДоговорПоСроку1 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку1();
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку1.Код = "12145";
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку1.Код = "12145";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку1[] _ЗадолжВидЗадолжКонтрагентДоговорПоСроку1 = { ЗадолжВидЗадолжКонтрагентДоговорПоСроку1 };
                    //ЗадолжВидЗфдолж1Договор.ПоСроку = _ЗадолжВидЗадолжКонтрагентДоговорПоСроку1;
                    //#endregion Договор

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор1[] _ЗадолжВидЗфдолж1Договор = { ЗадолжВидЗфдолж1Договор };
                    //ЗадолжВидЗадолжКонтрагент1.Договор = _ЗадолжВидЗфдолж1Договор;

                    //#endregion

                    //// ФайлДокументДебДебЗадолжЗадолжВидЗадолжКонтрагент1[] _ЗадолжВидЗадолжКонтрагент1 = { ЗадолжВидЗадолжКонтрагент1 };
                    //ЗадолжВидЗадолж1.Item = ЗадолжВидЗадолжКонтрагент1;

                    //#endregion Задолж

                    //#region Итого
                    //var ЗадолжВидИтого1 = new ФайлДокументДебКредЗадолжЗадолжВидИтого1();
                    //ЗадолжВидИтого1.Общая = "12145";
                    //ЗадолжВидИтого1.Просроч = "12145";

                    //#region ПоСроку
                    //var ЗадолжВидИтогоПоСроку1 = new ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку1();
                    //ЗадолжВидИтогоПоСроку1.Код = "11145";
                    //ЗадолжВидИтогоПоСроку1.Задолж = "12145";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку1[] _ЗадолжВидИтогоПоСроку1 = { ЗадолжВидИтогоПоСроку1 };
                    //ЗадолжВидИтого1.ПоСроку = _ЗадолжВидИтогоПоСроку1;
                    //#endregion Итого

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолж1[] _ЗадолжВидЗадолж1 = { ЗадолжВидЗадолж1 };
                    //ДебКсрЗадолжВид.Задолж = _ЗадолжВидЗадолж1;
                    //// ФайлДокументДебДебЗадолжЗадолжВидИтого1[] _ЗадолжВидИтого1 = { ЗадолжВидИтого1 };
                    //ДебКсрЗадолжВид.Итого = ЗадолжВидИтого1;
                    //ФайлДокументДебКредЗадолжЗадолжВид1[] _ДебКсрЗадолжВид = { ДебКсрЗадолжВид };
                    //ДебКредЗадолж.ДебКср = _ДебКсрЗадолжВид;
                    //Processing.ReportMessage(" ДебКсрЗадолжВид - закончили");
                    #endregion Конец_ДебКсрЗадолжВид

                    #region КредДсрЗадолжВид
                    //Processing.ReportMessage(" КредДсрЗадолжВид");
                    //var КредДсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид2();
                    //КредДсрЗадолжВид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид2.Item1;

                    //#region Задолж
                    //var ЗадолжВидЗадолж2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж2();

                    //#region Контрагент
                    //var ЗадолжВидЗадолжКонтрагент2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент2();
                    //ЗадолжВидЗадолжКонтрагент2.Наим = "qwerty";
                    //ЗадолжВидЗадолжКонтрагент2.ИНН = "1234567890";
                    //ЗадолжВидЗадолжКонтрагент2.КПП = "123456789";
                    //ЗадолжВидЗадолжКонтрагент2.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗадолжКонтрагент2.Общая = "123";

                    //#region Договор
                    //var ЗадолжВидЗфдолж2Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор2();
                    //ЗадолжВидЗфдолж2Договор.Наим = "qwerty";
                    //ЗадолжВидЗфдолж2Договор.Номер = "qwerty";
                    //ЗадолжВидЗфдолж2Договор.Дата = "01.01.2019";
                    //ЗадолжВидЗфдолж2Договор.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗфдолж2Договор.Общая = "12345";
                    //ЗадолжВидЗфдолж2Договор.Просроч = "12345";

                    //#region ПоСроку
                    //var ЗадолжВидЗадолжКонтрагентДоговорПоСроку2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку2();
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку2.Код = "12345";
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку2.Код = "12345";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку2[] _ЗадолжВидЗадолжКонтрагентДоговорПоСроку2 = { ЗадолжВидЗадолжКонтрагентДоговорПоСроку2 };
                    //ЗадолжВидЗфдолж2Договор.ПоСроку = _ЗадолжВидЗадолжКонтрагентДоговорПоСроку2;
                    //#endregion Договор

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор2[] _ЗадолжВидЗфдолж2Договор = { ЗадолжВидЗфдолж2Договор };
                    //ЗадолжВидЗадолжКонтрагент2.Договор = _ЗадолжВидЗфдолж2Договор;

                    //#endregion

                    //// ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент2[] _ЗадолжВидЗадолжКонтрагент2 = { ЗадолжВидЗадолжКонтрагент2 };
                    //ЗадолжВидЗадолж2.Контрагент = ЗадолжВидЗадолжКонтрагент2;

                    //#endregion Задолж

                    //#region Итого
                    //var ЗадолжВидИтого2 = new ФайлДокументДебКредЗадолжЗадолжВидИтого2();
                    //ЗадолжВидИтого2.Общая = "12345";
                    //ЗадолжВидИтого2.Просроч = "12345";

                    //#region ПоСроку
                    //var ЗадолжВидИтогоПоСроку2 = new ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку2();
                    //ЗадолжВидИтогоПоСроку2.Код = "12345";
                    //ЗадолжВидИтогоПоСроку2.Задолж = "12345";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку2[] _ЗадолжВидИтогоПоСроку2 = { ЗадолжВидИтогоПоСроку2 };
                    //ЗадолжВидИтого2.ПоСроку = _ЗадолжВидИтогоПоСроку2;
                    //#endregion Итого

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолж2[] _ЗадолжВидЗадолж2 = { ЗадолжВидЗадолж2 };
                    //КредДсрЗадолжВид.Задолж = _ЗадолжВидЗадолж2;
                    //// ФайлДокументДебКредЗадолжЗадолжВидИтого2[] _ЗадолжВидИтого2 = { ЗадолжВидИтого2 };
                    //КредДсрЗадолжВид.Итого = ЗадолжВидИтого2;
                    //ФайлДокументДебКредЗадолжЗадолжВид2[] _КредДсрЗадолжВид = { КредДсрЗадолжВид };
                    //ДебКредЗадолж.КредДср = _КредДсрЗадолжВид;
                    //Processing.ReportMessage(" КредДсрЗадолжВид - закончили");
                    #endregion Конец_КредДсрЗадолжВид

                    #region КредКсрЗадолжВид
                    if (Document.DefinitionName == "3_42_1_Расшифровка кредиторской задолженности")
                    {
                        Processing.ReportMessage(" КредКсрЗадолжВид");
                        var КредКсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид3();
                        КредКсрЗадолжВид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item1;
                        #region Задолж
                        Processing.ReportMessage(" Задолж");
                        List<string> notprochee = new List<string>() { "Прочее", "Прочие", "Итого" };
                        List<int> AccsRows = GetFields.GetTableParagraphWhereCellsNotMachText(Document, Processing, "Table", "Name", notprochee);
                        Processing.ReportMessage("AccsRows = " + AccsRows.Count);
                        Regex rgx = new Regex(@".*(Прочие|Прочее).*", RegexOptions.IgnoreCase);
                        List<int> ProchRows = GetFields.GetTableRowsWhereRegex(Document, Processing, "Table", "Name", rgx, 0, Document.Sections[0].Field("Table").Rows.Count);
                        Processing.ReportMessage("ProchRows = " + ProchRows.Count);
                        List<int> AllRows = new List<int>() { };
                        foreach (int it in AccsRows)
                            AllRows.Add(it);
                        foreach (int it in ProchRows)
                            AllRows.Add(it);

                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] ЗадолжВидЗадолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[AllRows.Count]; // Массив Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3[] ЗадолжВидЗадолжКонтрагент3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3[AccsRows.Count]; // массив Контрагентов
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолжПрочее1[] ЗадолжВидПрочее = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжПрочее1[ProchRows.Count]; //массив Прочее

                        Processing.ReportMessage("Контрагенты");
                        for (var l = 0; l < AccsRows.Count; l++) //цикл по Контрагентам
                        {
                            var _ЗадолжВидЗадолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            var _ЗадолжКонтрагент3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3();
                            #region данные
                            Processing.ReportMessage("Обрабатываем строку - " + l);
                            if (!String.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "Name", AccsRows[l]))) { _ЗадолжКонтрагент3.Наим = GetFields.GetCellByIndex(Document, Processing, "Table", "Name", AccsRows[l]).ToString(); } else { _ЗадолжКонтрагент3.Наим = "-"; };
                            if (!String.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "INN", AccsRows[l]))) { _ЗадолжКонтрагент3.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table", "INN", AccsRows[l]).ToString(); } else { };
                            //ЗадолжВидЗадолжКонтрагент3.КПП = "123456789";
                            //ЗадолжВидЗадолжКонтрагент3.ДатаВозн = "01.01.2019";
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "Sum", AccsRows[l]), out decimal a);
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "SumOverdue", AccsRows[l]), out decimal b);
                            // _ЗадолжКонтрагент3.Общая = a.ToString();
                            #endregion данные

                            #region Договор

                            Processing.ReportMessage("Договор");

                            var ЗадолжВидЗфдолж3Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3();
                            //ЗадолжВидЗфдолж3Договор.Наим = "qwerty";
                            //ЗадолжВидЗфдолж3Договор.Номер = "qwerty";
                            //ЗадолжВидЗфдолж3Договор.Дата = "01.01.2019";
                            //ЗадолжВидЗфдолж3Договор.ДатаВозн = "01.01.2019";
                            ЗадолжВидЗфдолж3Договор.Общая = a;//Decimal.ToInt32(a).ToString();
                            ЗадолжВидЗфдолж3Договор.Просроч = b;//Decimal.ToInt32(b).ToString();
                            //ЗадолжВидЗфдолж3Договор.Просроч = "12345";

                            #region ПоСроку
                            //var ЗадолжВидЗадолжКонтрагентДоговорПоСроку3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку3();
                            //ЗадолжВидЗадолжКонтрагентДоговорПоСроку3.Код = "12345";
                            //ЗадолжВидЗадолжКонтрагентДоговорПоСроку3.Код = "12345";
                            #endregion ПоСроку

                            //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку3[] _ЗадолжВидЗадолжКонтрагентДоговорПоСроку3 = { ЗадолжВидЗадолжКонтрагентДоговорПоСроку3 };
                            //ЗадолжВидЗфдолж3Договор.ПоСроку = _ЗадолжВидЗадолжКонтрагентДоговорПоСроку3;

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _ЗадолжВидЗфдолж3Договор = { ЗадолжВидЗфдолж3Договор };
                            _ЗадолжКонтрагент3.Договор = _ЗадолжВидЗфдолж3Договор;

                            #endregion Договор

                            ЗадолжВидЗадолжКонтрагент3[l] = _ЗадолжКонтрагент3;
                            _ЗадолжВидЗадолж3.Item = ЗадолжВидЗадолжКонтрагент3[l];
                            ЗадолжВидЗадолж3[l] = _ЗадолжВидЗадолж3;
                            КредКсрЗадолжВид.Задолж = ЗадолжВидЗадолж3;
                        }


                        for (var h = 0; h < ProchRows.Count; h++) //цикл по прочее
                        {
                            var _ЗадолжВидЗадолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            var _ЗадолжПрочее = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжПрочее1();
                            #region данные
                            if (!String.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table", "Name", ProchRows[h]))) { _ЗадолжПрочее.Наим = GetFields.GetCellByIndex(Document, Processing, "Table", "Name", ProchRows[h]).ToString(); } else { _ЗадолжПрочее.Наим = "-"; };
                            //  ЗадолжВидПрочее[l].Наим = GetFields.GetCellByIndex(Document, Processing, "Table", "Name", ProchRows[l]);
                            //ЗадолжВидПрочее.ДатаВозн = "";
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table", "Sum", ProchRows[h]), out decimal a);
                            _ЗадолжПрочее.Общая = a;//Decimal.ToInt32(a).ToString();
                            #endregion данные

                            ЗадолжВидПрочее[h] = _ЗадолжПрочее;
                            _ЗадолжВидЗадолж3.Item = ЗадолжВидПрочее[h];
                            ЗадолжВидЗадолж3[AccsRows.Count + h] = _ЗадолжВидЗадолж3;
                            КредКсрЗадолжВид.Задолж = ЗадолжВидЗадолж3;
                        }
                        Processing.ReportMessage("закрыли ЗадолжВидЗадолж3");
                        #endregion Задолж


                        ФайлДокументДебКредЗадолжЗадолжВид3[] _КредКсрЗадолжВид = { КредКсрЗадолжВид };
                        ДебКредЗадолж.КредКср = _КредКсрЗадолжВид;
                        Processing.ReportMessage("закрыли КредКср");


                    }
                    #endregion КредКсрЗадолжВид


                    Документ.ДебКредЗадолж = ДебКредЗадолж;
                    Processing.ReportMessage("ДебКредЗадолж - Закончили");

                }
                #endregion Документ_РасшДебКредЗад

                #region Документ_РасшДебКредЗад_СТАНДАРТНАЯ
                if (Document.DefinitionName == "3_42_1_Расшифровка кредиторской задолженности(Стандарт)" || Document.DefinitionName == "3_42_2_Расшифровка дебиторской задолженности(Стандарт)")
                {
                    Processing.ReportMessage("Начинаю формирование ХМЛ по документу - РасшДебКредЗад-СТАНДАРТ");
                    var ДебКредЗадолж = new ФайлДокументДебКредЗадолж();
                    ДебКредЗадолж.Период = PeriodQuarterType(GetFields.GetField(Document, Processing, "Date"));
                    
                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Date")) && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out var dateValue2)) { ДебКредЗадолж.ОтчетГод = dateValue2.Year.ToString(); } else { ДебКредЗадолж.ОтчетГод = DateTime.Now.Year.ToString(); }


                    ДебКредЗадолж.ОКЕИ = ФайлДокументДебКредЗадолжОКЕИ.Item384;

                    #region Интервал
                    ////тут должен быть цикл 
                    //var i = 0;
                    //var Интервал = new ФайлДокументДебКредЗадолжИнтервалыИнтервал();
                    //Интервал.Код = "000000";
                    //Интервал.Начало = "-";
                    //Интервал.Окончание = "-";
                    //// ФайлДокументДебКредЗадолжИнтервалыИнтервал[][] _Интервал = { Интервал } ;
                    ////ДебКредЗадолж.Интервалы = _Интервал;


                    #endregion Интервал

                    #region ДебДсрЗадолжВид
                    Processing.ReportMessage(" ДебДсрЗадолжВид");
                  
                    Processing.ReportMessage(" ДебДсрЗадолжВид - закончили");
                    #endregion Конец_ДебДсрЗадолжВид

                    #region ДебКсрЗадолжВид
                    if (Document.DefinitionName == "3_42_2_Расшифровка дебиторской задолженности(Стандарт)")
                    {
                        Processing.ReportMessage(" ДебКсрЗадолжВид");
                        ФайлДокументДебКредЗадолжЗадолжВид1[] ЗадолжВид1 = new ФайлДокументДебКредЗадолжЗадолжВид1[4]; //открываем ДебКср
                        log("открыли ДебКср");
                        #region Вид 1 Расчеты с покупателями/заказчиками
                        var _Вид = new ФайлДокументДебКредЗадолжЗадолжВид1(); //Вид
                        _Вид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид1.Item1;
                        #region Задолж
                     //   ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("TableShortSettlementsWithBuyersCustomers").Rows.Count];//Задолж
                    
                        

                     //   _Вид.Задолж = Задолж;

                        #endregion Задолж
                        ЗадолжВид1[0] = _Вид;
                        #endregion Вид 1

                        #region Вид 2  Авансы выданные
                        #endregion Вид 2 

                        #region Вид 4 Задолженность учредителей
                        #endregion Вид 4

                        #region Вид 6 Прочая
                        #endregion Вид 6

                        ДебКредЗадолж.ДебКср = ЗадолжВид1; //Закрываем ДебКср
                        Processing.ReportMessage("закрыли ДебКср");
                    }
                   
                    #endregion Конец_ДебКсрЗадолжВид

                    #region КредДсрЗадолжВид
                    //Processing.ReportMessage(" КредДсрЗадолжВид");
                    //var КредДсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид2();
                    //КредДсрЗадолжВид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид2.Item1;

                    //#region Задолж
                    //var ЗадолжВидЗадолж2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж2();

                    //#region Контрагент
                    //var ЗадолжВидЗадолжКонтрагент2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент2();
                    //ЗадолжВидЗадолжКонтрагент2.Наим = "qwerty";
                    //ЗадолжВидЗадолжКонтрагент2.ИНН = "1234567890";
                    //ЗадолжВидЗадолжКонтрагент2.КПП = "123456789";
                    //ЗадолжВидЗадолжКонтрагент2.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗадолжКонтрагент2.Общая = "123";

                    //#region Договор
                    //var ЗадолжВидЗфдолж2Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор2();
                    //ЗадолжВидЗфдолж2Договор.Наим = "qwerty";
                    //ЗадолжВидЗфдолж2Договор.Номер = "qwerty";
                    //ЗадолжВидЗфдолж2Договор.Дата = "01.01.2019";
                    //ЗадолжВидЗфдолж2Договор.ДатаВозн = "01.01.2019";
                    //ЗадолжВидЗфдолж2Договор.Общая = "12345";
                    //ЗадолжВидЗфдолж2Договор.Просроч = "12345";

                    //#region ПоСроку
                    //var ЗадолжВидЗадолжКонтрагентДоговорПоСроку2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку2();
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку2.Код = "12345";
                    //ЗадолжВидЗадолжКонтрагентДоговорПоСроку2.Код = "12345";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговорПоСроку2[] _ЗадолжВидЗадолжКонтрагентДоговорПоСроку2 = { ЗадолжВидЗадолжКонтрагентДоговорПоСроку2 };
                    //ЗадолжВидЗфдолж2Договор.ПоСроку = _ЗадолжВидЗадолжКонтрагентДоговорПоСроку2;
                    //#endregion Договор

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор2[] _ЗадолжВидЗфдолж2Договор = { ЗадолжВидЗфдолж2Договор };
                    //ЗадолжВидЗадолжКонтрагент2.Договор = _ЗадолжВидЗфдолж2Договор;

                    //#endregion

                    //// ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент2[] _ЗадолжВидЗадолжКонтрагент2 = { ЗадолжВидЗадолжКонтрагент2 };
                    //ЗадолжВидЗадолж2.Контрагент = ЗадолжВидЗадолжКонтрагент2;

                    //#endregion Задолж

                    //#region Итого
                    //var ЗадолжВидИтого2 = new ФайлДокументДебКредЗадолжЗадолжВидИтого2();
                    //ЗадолжВидИтого2.Общая = "12345";
                    //ЗадолжВидИтого2.Просроч = "12345";

                    //#region ПоСроку
                    //var ЗадолжВидИтогоПоСроку2 = new ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку2();
                    //ЗадолжВидИтогоПоСроку2.Код = "12345";
                    //ЗадолжВидИтогоПоСроку2.Задолж = "12345";
                    //#endregion ПоСроку

                    //ФайлДокументДебКредЗадолжЗадолжВидИтогоПоСроку2[] _ЗадолжВидИтогоПоСроку2 = { ЗадолжВидИтогоПоСроку2 };
                    //ЗадолжВидИтого2.ПоСроку = _ЗадолжВидИтогоПоСроку2;
                    //#endregion Итого

                    //ФайлДокументДебКредЗадолжЗадолжВидЗадолж2[] _ЗадолжВидЗадолж2 = { ЗадолжВидЗадолж2 };
                    //КредДсрЗадолжВид.Задолж = _ЗадолжВидЗадолж2;
                    //// ФайлДокументДебКредЗадолжЗадолжВидИтого2[] _ЗадолжВидИтого2 = { ЗадолжВидИтого2 };
                    //КредДсрЗадолжВид.Итого = ЗадолжВидИтого2;
                    //ФайлДокументДебКредЗадолжЗадолжВид2[] _КредДсрЗадолжВид = { КредДсрЗадолжВид };
                    //ДебКредЗадолж.КредДср = _КредДсрЗадолжВид;
                    //Processing.ReportMessage(" КредДсрЗадолжВид - закончили");
                    #endregion Конец_КредДсрЗадолжВид

                    #region КредКсрЗадолжВид
                    if (Document.DefinitionName == "3_42_1_Расшифровка кредиторской задолженности(Стандарт)")
                    {

                       
                        Processing.ReportMessage(" КредКсрЗадолжВид");
                       // ФайлДокументДебКредЗадолжЗадолжВид3[] КредКсрЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид3[1]; //КредКср
                        log("Объявили КредКсрЗадолжВид");
                        ФайлДокументДебКредЗадолжЗадолжВид3[] ЗадолжВид = new ФайлДокументДебКредЗадолжЗадолжВид3[6]; //ЗадолжВид

                        #region ВИД 1

                        log("Ровняем вид");
                        var _Вид = new ФайлДокументДебКредЗадолжЗадолжВид3(); //Вид
                        _Вид.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item1;
                       
                        #region Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("Table2").Rows.Count];//Задолж
                        log("Начинаем цикл по таблице");
                        for (int rows = 0; rows < Document.Sections[0].Field("Table2").Rows.Count; rows++)
                        {
                            #region контрагент
                            var Контрагент = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3(); //Контрагент
                            log("Обрабатываем строку " + rows);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table2", "AgentName", rows))) { Контрагент.Наим = GetFields.GetCellByIndex(Document, Processing, "Table2", "AgentName", rows); } else { Контрагент.Наим = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table2", "INN", rows))) {Контрагент.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table2", "INN", rows); } else { }
                            log("Записали Контрагента и ИНН");

                            #region Договор
                            var Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3(); //Договор
                             log("Зашли в договор");
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table2", "Date", rows)) && DateTime.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table2", "Date", rows), out var dtResult)) {Договор.Дата = dtResult.ToString("dd.MM.yyyy"); } else { /*Договор.Дата = "";*/ }
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table2", "Sum", rows).Replace(".", k()).Replace(",", k()), out decimal a);
                            Договор.Общая = a; 
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table2", "SumF", rows).Replace(".", k()).Replace(",", k()), out decimal b);
                            Договор.Просроч = b; 
                           
                            log("Записали атрибуты договора");
                      
                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _Договор = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] { Договор };
                            Контрагент.Договор = _Договор;
                            log("закрыли договор");
                            #endregion Договор
                            var _Задолж= new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            _Задолж.Item = Контрагент;
                            Задолж[rows] = _Задолж;
                            log("Приравняли к строке Задолж");
                            #endregion контрагент
                        }

                        _Вид.Задолж = Задолж;
                        ЗадолжВид[0] = _Вид;

                        log("Заровняли Задолж");
                        #endregion Задолж

                        #endregion ВИД 1

                        #region ВИД 2

                        log("Ровняем вид 2");
                        var _Вид2 = new ФайлДокументДебКредЗадолжЗадолжВид3(); //Вид
                        _Вид2.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item2;

                        #region Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("Table2").Rows.Count];//Задолж
                        log("Начинаем цикл по таблице"+ "  - Document.Sections[0].Field(Table2).Rows.Count = "+ Document.Sections[0].Field("Table2").Rows.Count);

                        for (int rows1 = 0; rows1 < Document.Sections[0].Field("Table2").Rows.Count; rows1++)
                        {
                            #region контрагент
                            var Контрагент2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3(); //Контрагент
                            log("Обрабатываем строку " + rows1);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table2", "AgentName", rows1))) { Контрагент2.Наим = GetFields.GetCellByIndex(Document, Processing, "Table2", "AgentName", rows1); } else { Контрагент2.Наим = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table2", "INN", rows1))) { Контрагент2.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table2", "INN", rows1); } else { }
                            log("Записали Контрагента и ИНН");

                            #region Договор
                            var Договор2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3(); //Договор
                            log("Зашли в договор");
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table2", "Date", rows1)) && DateTime.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table2", "Date", rows1), out var dtResult)) { Договор2.Дата = dtResult.ToString("dd.MM.yyyy"); } else {/* Договор2.Дата = "";*/ }
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table2", "Sum", rows1).Replace(".", k()).Replace(",", k()), out decimal a);
                            Договор2.Общая = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table2", "SumF", rows1).Replace(".", k()).Replace(",", k()), out decimal b);
                            Договор2.Просроч = b;

                            log("Записали атрибуты договора");

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _Договор2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] { Договор2 };
                            Контрагент2.Договор = _Договор2;
                            log("закрыли договор");
                            #endregion Договор
                            var _Задолж2 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            log("var _Задолж2");
                            _Задолж2.Item = Контрагент2;
                            log("Заровняли Контрагент2" + "; rows1 = "+ rows1);
                            Задолж2[rows1] = _Задолж2;
                            log("Приравняли к строке Задолж");
                            #endregion контрагент
                        }

                        _Вид2.Задолж = Задолж2;
                        ЗадолжВид[1] = _Вид2;

                        log("Заровняли Задолж");
                        #endregion Задолж

                        #endregion ВИД 2

                        #region ВИД 3

                        log("Ровняем вид 3");
                        var _Вид3 = new ФайлДокументДебКредЗадолжЗадолжВид3(); //Вид
                        _Вид3.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item3;

                        #region Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("Table2").Rows.Count];//Задолж
                        log("Начинаем цикл по таблице");
                        for (int rows = 0; rows < Document.Sections[0].Field("Table4").Rows.Count; rows++)
                        {
                            #region контрагент
                            var Контрагент3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3(); //Контрагент
                            log("Обрабатываем строку " + rows);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table4", "AgentName", rows))) { Контрагент3.Наим = GetFields.GetCellByIndex(Document, Processing, "Table4", "AgentName", rows); } else { Контрагент3.Наим = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table4", "INN", rows))) { Контрагент3.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table4", "INN", rows); } else { }
                            log("Записали Контрагента и ИНН");

                            #region Договор
                            var Договор3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3(); //Договор
                            log("Зашли в договор");
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table4", "Date", rows)) && DateTime.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table4", "Date", rows), out var dtResult)) { Договор3.Дата = dtResult.ToString("dd.MM.yyyy"); } else { /*Договор3.Дата = "";*/ }
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table4", "Sum", rows).Replace(".", k()).Replace(",", k()), out decimal a);
                            Договор3.Общая = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table4", "SumF", rows).Replace(".", k()).Replace(",", k()), out decimal b);
                            Договор3.Просроч = b;

                            log("Записали атрибуты договора");

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _Договор3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] { Договор3 };
                            Контрагент3.Договор = _Договор3;
                            log("закрыли договор");
                            #endregion Договор
                            var _Задолж3 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            _Задолж3.Item = Контрагент3;
                            Задолж3[rows] = _Задолж3;
                            log("Приравняли к строке Задолж");
                            #endregion контрагент
                        }

                        _Вид3.Задолж = Задолж3;
                        ЗадолжВид[2] = _Вид3;

                        log("Заровняли Задолж");
                        #endregion Задолж

                        #endregion ВИД 3

                        #region ВИД 4

                        log("Ровняем вид 3");
                        var _Вид4 = new ФайлДокументДебКредЗадолжЗадолжВид3(); //Вид
                        _Вид4.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item4;

                        #region Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж4 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("Table2").Rows.Count];//Задолж
                        log("Начинаем цикл по таблице");
                        for (int rows = 0; rows < Document.Sections[0].Field("Table3").Rows.Count; rows++)
                        {
                            #region контрагент
                            var Контрагент4 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3(); //Контрагент
                            log("Обрабатываем строку " + rows);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table3", "AgentName", rows))) { Контрагент4.Наим = GetFields.GetCellByIndex(Document, Processing, "Table3", "AgentName", rows); } else { Контрагент4.Наим = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table3", "INN", rows))) { Контрагент4.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table3", "INN", rows); } else { }
                            log("Записали Контрагента и ИНН");

                            #region Договор
                            var Договор4 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3(); //Договор
                            log("Зашли в договор");
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table3", "Date", rows)) && DateTime.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table3", "Date", rows), out var dtResult)) { Договор4.Дата = dtResult.ToString("dd.MM.yyyy"); } else { /*Договор4.Дата = "";*/ }
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table3", "Sum", rows).Replace(".", k()).Replace(",", k()), out decimal a);
                            Договор4.Общая = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table3", "SumF", rows).Replace(".", k()).Replace(",", k()), out decimal b);
                            Договор4.Просроч = b;

                            log("Записали атрибуты договора");

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _Договор4 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] { Договор4 };
                            Контрагент4.Договор = _Договор4;
                            log("закрыли договор");
                            #endregion Договор
                            var _Задолж4 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            _Задолж4.Item = Контрагент4;
                            Задолж4[rows] = _Задолж4;
                            log("Приравняли к строке Задолж");
                            #endregion контрагент
                        }

                        _Вид4.Задолж = Задолж4;
                        ЗадолжВид[3] = _Вид4;

                        log("Заровняли Задолж");
                        #endregion Задолж

                        #endregion ВИД 4

                        #region ВИД 5

                        log("Ровняем вид 3");
                        var _Вид5 = new ФайлДокументДебКредЗадолжЗадолжВид3(); //Вид
                        _Вид5.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item5;

                        #region Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж5 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("Table2").Rows.Count];//Задолж
                        log("Начинаем цикл по таблице");
                        for (int rows = 0; rows < Document.Sections[0].Field("Table5").Rows.Count; rows++)
                        {
                            #region контрагент
                            var Контрагент5 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3(); //Контрагент
                            log("Обрабатываем строку " + rows);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table5", "AgentName", rows))) { Контрагент5.Наим = GetFields.GetCellByIndex(Document, Processing, "Table5", "AgentName", rows); } else { Контрагент5.Наим = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table5", "INN", rows))) { Контрагент5.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table5", "INN", rows); } else { }
                            log("Записали Контрагента и ИНН");

                            #region Договор
                            var Договор5 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3(); //Договор
                            log("Зашли в договор");
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table5", "Date", rows)) && DateTime.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table5", "Date", rows), out var dtResult)) { Договор5.Дата = dtResult.ToString("dd.MM.yyyy"); } else {/* Договор5.Дата = ""; */}
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table5", "Sum", rows).Replace(".", k()).Replace(",", k()), out decimal a);
                            Договор5.Общая = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table5", "SumF", rows).Replace(".", k()).Replace(",", k()), out decimal b);
                            Договор5.Просроч = b;

                            log("Записали атрибуты договора");

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _Договор5 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] { Договор5 };
                            Контрагент5.Договор = _Договор5;
                            log("закрыли договор");
                            #endregion Договор
                            var _Задолж5 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            _Задолж5.Item = Контрагент5;
                            Задолж5[rows] = _Задолж5;
                            log("Приравняли к строке Задолж");
                            #endregion контрагент
                        }

                        _Вид5.Задолж = Задолж5;
                        ЗадолжВид[4] = _Вид5;

                        log("Заровняли Задолж");
                        #endregion Задолж

                        #endregion ВИД 5

                        #region ВИД 6

                        log("Ровняем вид 3");
                        var _Вид6 = new ФайлДокументДебКредЗадолжЗадолжВид3(); //Вид
                        _Вид6.Вид = ФайлДокументДебКредЗадолжЗадолжВидВид3.Item6;

                        #region Задолж
                        ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[] Задолж6 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3[Document.Sections[0].Field("Table2").Rows.Count];//Задолж
                        log("Начинаем цикл по таблице");
                        for (int rows = 0; rows < Document.Sections[0].Field("Table6").Rows.Count; rows++)
                        {
                            #region контрагент
                            var Контрагент6 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагент3(); //Контрагент
                            log("Обрабатываем строку " + rows);
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table6", "AgentName", rows))) { Контрагент6.Наим = GetFields.GetCellByIndex(Document, Processing, "Table6", "AgentName", rows); } else { Контрагент6.Наим = "-"; }
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table6", "INN", rows))) { Контрагент6.ИНН = GetFields.GetCellByIndex(Document, Processing, "Table6", "INN", rows); } else { }
                            log("Записали Контрагента и ИНН");

                            #region Договор
                            var Договор6 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3(); //Договор
                            log("Зашли в договор");
                            if (!string.IsNullOrEmpty(GetFields.GetCellByIndex(Document, Processing, "Table6", "Date", rows)) && DateTime.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table6", "Date", rows), out var dtResult)) { Договор6.Дата = dtResult.ToString("dd.MM.yyyy"); } else {/* Договор6.Дата = "";*/ }
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table6", "Sum", rows).Replace(".", k()).Replace(",", k()), out decimal a);
                            Договор6.Общая = a;
                            Decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Table6", "SumF", rows).Replace(".", k()).Replace(",", k()), out decimal b);
                            Договор6.Просроч = b;

                            log("Записали атрибуты договора");

                            ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] _Договор6 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолжКонтрагентДоговор3[] { Договор6 };
                            Контрагент6.Договор = _Договор6;
                            log("закрыли договор");
                            #endregion Договор
                            var _Задолж6 = new ФайлДокументДебКредЗадолжЗадолжВидЗадолж3();
                            _Задолж6.Item = Контрагент6;
                            Задолж6[rows] = _Задолж6;
                            log("Приравняли к строке Задолж");
                            #endregion контрагент
                        }

                        _Вид6.Задолж = Задолж6;
                        ЗадолжВид[5] = _Вид6;

                        log("Заровняли Задолж");
                        #endregion Задолж

                        #endregion ВИД 6

                        ДебКредЗадолж.КредКср = ЗадолжВид;
                          
                        Processing.ReportMessage("закрыли КредКср");


                    }
                    #endregion КредКсрЗадолжВид


                    Документ.ДебКредЗадолж = ДебКредЗадолж;
                    Processing.ReportMessage("ДебКредЗадолж - Закончили");

                }
                #endregion Документ_РасшДебКредЗад_СТАНДАРТ

                #region Документ_прочДохРасх

                if (Document.DefinitionName == "3_70_Расшифровка прочих доходов расходов")
                {
                    Processing.ReportMessage("3_70_Расшифровка прочих доходов расходов");
                    //logger.Trace("Сериализация 3_70_Расшифровка прочих доходов расходов");
                    var ПрочДохРасх = new ФайлДокументПрочДохРасх();
                    ПрочДохРасх.Период = PeriodQuarterType(GetFields.GetField(Document, Processing, "Date"));


                    //logger.Debug("ПрочДохРасхПериод = "+ ПрочДохРасхПериод.ToString());

                    if (!string.IsNullOrEmpty(GetFields.GetField(Document, Processing, "Date")) && DateTime.TryParse(GetFields.GetField(Document, Processing, "Date"), out var dateValue2)) { ПрочДохРасх.ОтчетГод = dateValue2.Year.ToString(); } else { ПрочДохРасх.ОтчетГод = DateTime.Now.Year.ToString(); }

                    //ПрочДохРасх.ОтчетГод = "";
                    //logger.Debug("ПрочДохРасх.ОтчетГод = "+ ПрочДохРасх.ОтчетГод.ToString());
                    var ПрочДохРасхОКЕИ = new ФайлДокументПрочДохРасхОКЕИ();
                    ПрочДохРасх.ОКЕИ = ФайлДокументПрочДохРасхОКЕИ.Item384;

                    #region Доходы

                    //logger.Trace("Обрабатываем доходы");
                    var ПрочДохРасхДоходы = new ФайлДокументПрочДохРасхДоходы();
                    ПрочДохРасхДоходы.Итого = GetFields.GetCellWhere(Document, Processing, "Dohod", "Итого", "Name", "Sum").Replace(",", k()).Replace(".", k());
                    //logger.Debug("ПрочДохРасхДоходы.Итого = " + ПрочДохРасхДоходы.Итого.ToString());
                    ФайлДокументПрочДохРасхДоходыПоСтатье[] _ПрочДохРасхДоходыПоСтатье =
                        new ФайлДокументПрочДохРасхДоходыПоСтатье[Document.Sections[0].Field("Dohod").Rows.Count];
                    for (int par = 0; par < Document.Sections[0].Field("Dohod").Rows.Count; par++)
                    {
                        if (!GetFields.GetCellByIndex(Document, Processing, "Dohod", "Name", par).ToLower()
                            .Contains("итого"))
                        {
                            //logger.Trace("обрабатываем строку таблицы № " + par);
                            var ПрочДохРасхДоходыПоСтатье = new ФайлДокументПрочДохРасхДоходыПоСтатье();
                            ПрочДохРасхДоходыПоСтатье.Наим =
                                GetFields.GetCellByIndex(Document, Processing, "Dohod", "Name", par);
                            //logger.Debug("ПрочДохРасхДоходыПоСтатье.Наим " + ПрочДохРасхДоходыПоСтатье.Наим.ToString());
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Dohod", "Sum", par),
                                out decimal a);
                            ПрочДохРасхДоходыПоСтатье.Сумма = a; //Decimal.ToInt32(a).ToString();
                            //logger.Debug("ПрочДохРасхДоходыПоСтатье.Наим " + ПрочДохРасхДоходыПоСтатье.Наим.ToString());
                            _ПрочДохРасхДоходыПоСтатье[par] = ПрочДохРасхДоходыПоСтатье;
                            ПрочДохРасхДоходы.ПоСтатье = _ПрочДохРасхДоходыПоСтатье;
                            //logger.Trace("Закончили доходы");
                        }
                    }

                    ПрочДохРасх.Доходы = ПрочДохРасхДоходы;

                    #endregion Доходы

                    #region Расходы

                    //logger.Trace("Обрабатываем расходы");
                    var ПрочДохРасхРасходы = new ФайлДокументПрочДохРасхРасходы();
                    ФайлДокументПрочДохРасхРасходыПоСтатье[] _ПрочДохРасхРасходыПоСтатье =new ФайлДокументПрочДохРасхРасходыПоСтатье[Document.Sections[0].Field("Rashod").Rows.Count];
                    ПрочДохРасхРасходы.Итого =GetFields.GetCellWhere(Document, Processing, "Rashod", "Итого", "Name", "Sum").Replace(",", k()).Replace(".", k());
                    //logger.Debug(" ПрочДохРасхРасходы.Итого = " + ПрочДохРасхРасходы.Итого.ToString());

                    for (int par = 0; par < Document.Sections[0].Field("Rashod").Rows.Count; par++)
                    {
                        if (!GetFields.GetCellByIndex(Document, Processing, "Rashod", "Name", par).ToLower().Contains("итого"))
                        {
                            //logger.Trace("обрабатываем строку таблицы № "+par);
                            var ПрочДохРасхРасходыПоСтатье = new ФайлДокументПрочДохРасхРасходыПоСтатье();
                            ПрочДохРасхРасходыПоСтатье.Наим =
                                GetFields.GetCellByIndex(Document, Processing, "Rashod", "Name", par);
                            //logger.Debug("ПрочДохРасхРасходыПоСтатье.Наим = "+ ПрочДохРасхРасходыПоСтатье.Наим.ToString());
                            decimal.TryParse(GetFields.GetCellByIndex(Document, Processing, "Rashod", "Sum", par),
                                out decimal b);
                            ПрочДохРасхРасходыПоСтатье.Сумма = b; // Decimal.ToInt32(b).ToString();

                            //logger.Debug("ПрочДохРасхРасходыПоСтатье.Сумма = "+ ПрочДохРасхРасходыПоСтатье.Сумма.ToString());
                            _ПрочДохРасхРасходыПоСтатье[par] = ПрочДохРасхРасходыПоСтатье;
                            //ФайлДокументПрочДохРасхРасходыПоСтатье[] _ПрочДохРасхРасходыПоСтатье ={ПрочДохРасхРасходыПоСтатье};
                            ПрочДохРасхРасходы.ПоСтатье = _ПрочДохРасхРасходыПоСтатье;
                        }
                    }

                    ПрочДохРасх.Расходы = ПрочДохРасхРасходы;
                    //logger.Trace("Закончили расходы");

                    #endregion Расходы
                    Документ.ПрочДохРасх = ПрочДохРасх;
                }

                #endregion Документ_прочДохРасх

                #region КОНЦОВКА_XML
                // Исходное имя файла.

                string sourceFileName = @"";

                if (Document.Pages[0].ImageSourceType == @"File")
                {

                    sourceFileName = Path.GetFileName(Document.Pages[0].ImageSource);

                }
                else if (Document.Pages[0].ImageSourceType == @"Scanner" || Document.Pages[0].ImageSourceType == @"Custom")
                {

                    sourceFileName = Path.GetFileName(Document.Pages[0].ImageSourceFileSubPath);

                }

                Документ.ПутьДоОригинала = ExpPath.Replace(".xml", ".pdf");//Document.Batch.Project.ExportRootPath + @"\" + Document.Batch.Name + @"\" + sourceFileName;
                Документ.СвНП = СвНп;
                Документ.Подписант = Подписант;
                #endregion

                #endregion Документы

                Файл.Документ = Документ;

                #endregion Файл
                serializer.Serialize(stream, Файл);

            }
            catch (Exception er)
            {
                Processing.ReportError("Произошла ошибка сериализации XML: " + er.Message + @" информация об исключении: " + er.InnerException + @" Объект или сборка: " + er.Source + @" Метод, вызвавший исключение: " + er.TargetSite);

            }
            finally
            {
                stream.Close();

                ExportPDF(Document, Processing, iso, ExpPath.Replace(".xml", ".pdf"));
            }
        }

        /// <summary>
        /// Метод экспорта оригиналов изображений
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="ExpPath"></param>
        public static void ExportOriginals(IDocument Document, IProcessingCallback Processing, string ExpPath)
        {
            Processing.ReportMessage("Требуется экспорт оригиналов ");
            // Обходим страницы.
            for (int i = 0; i < Document.Pages.Count; i++)
            {
                IPage page = Document.Pages[i];
                // Проверяем наличие исходного файла.
                //Processing.ReportMessage("Проверяем наличие исходного файла страницы ");
                if (page.SourceFileGUID == string.Empty) { continue; }
                string sourceFileName = @"";
                if (page.ImageSourceType == @"File")
                {
                    sourceFileName = Path.GetFileName(page.ImageSource);
                }
                else if (page.ImageSourceType == @"Scanner" || page.ImageSourceType == @"Custom") { sourceFileName = Path.GetFileName(page.ImageSourceFileSubPath); }
                string sourceFilePath = ExpPath + @"\" + sourceFileName;
                // Проверяем, что этот файл ещё не сохраняли.
                Processing.ReportMessage("Путь экспорта: " + sourceFilePath);
                if (File.Exists(sourceFilePath)) { continue; }
                // Сохраняем файл.
                Processing.ReportMessage("Saving source file: " + sourceFileName);
                page.SaveSourceFile(sourceFilePath);

                // Document.Properties.Set("Exported", "true");
            }
        }

        /// <summary>
        /// Метод экспортирует PDF с текстовым слоем
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="iso"></param>
        /// <param name="ExpPath"></param>
        public static void ExportPDF(IDocument Document, IProcessingCallback Processing, IExportImageSavingOptions iso, string ExpPath)
        {
            try
            {
                Processing.ReportMessage("Экспортируем PDF");
                // Проверяем, что этот файл ещё не сохраняли.
                //TPdfTextSearchAreaType PTSAT_AllPages = new TPdfTextSearchAreaType();
                // TImageCompressionType ICT_Uncompressed = new TImageCompressionType() ;

                Processing.ReportMessage("Путь экспорта: " + ExpPath);
                iso.PdfTextSearchArea = TPdfTextSearchAreaType.PTSAT_AllPages;
                iso.Format = "pdf-s";
                iso.ColorType = "FullColor";
                iso.ImageCompressionType = TImageCompressionType.ICT_Uncompressed;
                iso.Quality = 100;
                iso.SaveAttachmentsToPdf = true;
                iso.ShouldOverwrite = true;
                Processing.ReportMessage("PDF-s ");
                iso.AddProperFileExt = true;
                iso.Resolution = 300;
                Document.SaveAs(ExpPath, iso);
            }
            catch (Exception er)
            {
                Processing.ReportError("Произошла ошибка экспорта PDF: " + er.Message + "; информация об исключении: " + er.InnerException + "; Объект или сборка: " + er.Source + "; Метод, вызвавший исключение: " + er.TargetSite);
            }

        }

        /// <summary>
        /// Сохраняет файл с заданым названием в папке экспорта
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="filename"></param>
        /// <param name="ExpPath"></param>
        public static void Puttxt(IDocument Document, IProcessingCallback Processing, string filename, string ExpPath = "")
        {

            if (Document.Batch.Properties.Has("Export") && !string.IsNullOrEmpty(Document.Batch.Properties.Get("Export")))
            {
                try
                {
                    Processing.ReportMessage("Path.GetDirectoryName(Document.Batch.Properties.Get(Export))" + Path.GetDirectoryName(Document.Batch.Properties.Get("Export")));
                    ExpPath = Path.GetDirectoryName(Document.Batch.Properties.Get("Export") + @"\" + Document.Batch.Name  + @"\");//+ "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))
                    Processing.ReportMessage("Путь экспорта из рег. параметра: " + ExpPath);
                }
                catch
                {
                    Processing.ReportWarning("В Рег.параметре пакета указан не корректный путь экспорта. Экспорт производится относительно папки импорта");
                    ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name  + @"\";//+ "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))
                }
            }
            else
            { ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name + @"\"; }//+ "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource)) 
            Processing.ReportMessage("Документ будет экспортироваться в " + ExpPath);

                //ExpPath = Path.GetDirectoryName(ExpPath);

            if (!Directory.Exists(ExpPath))
                Directory.CreateDirectory(ExpPath);
            string filepath = ExpPath + @"\" + filename + @".txt";

            if (!File.Exists(filepath))
            {
                using (var myfile = File.Create(filepath)) { myfile.Close(); };

            }



        }

        /// <summary>
        /// Сохраняет тхт файл с заданным именем.  
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="filename">имя файла</param>
        /// <param name="ExpPath">необязательный параметр по которому выгружать файл</param>
        public static void Puttxtstatus(IDocument Document, string filename, string ExpPath = "")
        {

            if (Document.Batch.Properties.Has("Export") && Document.Batch.Properties.Get("Export") != string.Empty)
            {
                try
                {
                    ExpPath = Document.Batch.Properties.Get("Export") + @"\" + Document.Batch.Name + @"\"; // + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))

                }
                catch
                {
                    ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name + @"\"; // + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))
                }
            }
            else
            { ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name + @"\"; }// + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))

            //ExpPath = Path.GetDirectoryName(ExpPath);

            if (!Directory.Exists(ExpPath))
                Directory.CreateDirectory(ExpPath);
            string filepath = ExpPath + @"\" + filename + @".txt";

            if (!File.Exists(filepath))
            {
                using (var myfile = File.Create(filepath)) { myfile.Close(); };

            }


        }

        /// <summary>
        /// Экспорт нераспознанных документов
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Processing"></param>
        /// <param name="ExpPath"></param>
        public static void ExportNonRecognisedOriginals(IDocument Document, IProcessingCallback Processing, string ExpPath = "")
        {
            Processing.ReportMessage("Экспортирую нераспознанные оригиналы");

            if (Document.Batch.Properties.Has("Export") && !string.IsNullOrEmpty(Document.Batch.Properties.Get("Export")))
            {
                try
                {
                    Processing.ReportMessage("Path.GetDirectoryName(Document.Batch.Properties.Get(Export))" + Path.GetDirectoryName(Document.Batch.Properties.Get("Export")));
                    ExpPath = Path.GetDirectoryName(Document.Batch.Properties.Get("Export") + @"\" + Document.Batch.Name + @"\"); // + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))
                    Processing.ReportMessage("Путь экспорта из рег. параметра: " + ExpPath);
                }
                catch
                {
                    Processing.ReportWarning("В Рег.параметре пакета указан не корректный путь экспорта. Экспорт производится относительно папки импорта");
                    ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name+ @"\";// + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource)) 
                }
            }
            else
            { ExpPath = Path.GetDirectoryName(Path.GetDirectoryName(Document.Pages[0].ImageSource)) + @"\Export" + @"\" + Document.Batch.Name + @"\"; } // + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(Document.Pages[0].ImageSource))
            Processing.ReportMessage("Документ будет экспортироваться в " + ExpPath);
            //ExpPath = Path.GetDirectoryName(ExpPath);

            if (!Directory.Exists(ExpPath))
                Directory.CreateDirectory(ExpPath);

            //ExportOriginals(Document, Processing, ExpPath);
            //Processing.ReportMessage("Оригиналы выгружены");

            Puttxtstatus(Document, "HasNotRecognisedDocs", ExpPath);
        }

    }
}
