using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ABBYY.FlexiCapture;
using ABBYY;
using ABBYY.FlexiCapture.ClientUI;


namespace StageTools
{
    public static partial class Forcematch
    {
        /// <summary>
        /// Метод проверяет есть ли указанный класс в списке Определений для наложения
        /// </summary>
        /// <param name="clas">класс</param>
        /// <param name="DefList">список ОД</param>
        /// <returns>true или false</returns>
        public static bool has(string clas, string[] DefList)
        {
            bool s = false;
            foreach (string i in DefList)
            {
                if (i.Contains(clas))
                    s = true;
            }
            return s;
        }

        /// <summary>
        ///метод выбирает только те ОД которые содержат искомый текст (например результирующий класс)
        /// </summary>
        /// <param name="clas">любой текст который будем искать в списке ОД</param>
        /// <param name="DefList">список ОД</param>
        /// <returns>массив ОД с совпадениями (без всяких разделителей)</returns>
        public static string[] IsReference(string clas, string[] DefList)
        {
            string[] RefList;
            string s = "";
            foreach (string i in DefList)
            {
                if (i.Contains(clas))
                    s += i + ";";
            }
            return RefList = s.Split(';');
        }

        /// <summary>
        /// Метод проверяет есть ли ОД с таким классом и накладывает его, а если такого нет, то накладывает Документ
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="PageIndex"></param>
        /// <param name="Matching"></param>
        public static void Beforematching(IDocument Document, int PageIndex, IMatchingInfo Matching)
        {
            string[] DefList = Matching.DefinitionsList.Split(';');
            string ResultDefList = string.Empty;
            if (Document.Pages[PageIndex].Comment.Length>1) //если класс определился
            {
                if (Document.Pages[PageIndex].Comment.Contains("Справка из банка") == true || has(Document.Pages[PageIndex].Comment, DefList) == true) // если класс содержит справку из банка или класс есть в списке ОД 
                {
                    if (Document.Pages[PageIndex].Comment.Contains("Справка из банка") == true && IsReference("Справка из банка", DefList).Length > 0) //если класс содержит справку из банка и Справка есть в списке ОД
                    {
                        foreach (string i in IsReference("Справка из банка", DefList)) // добаввляем все справки в ResultDefList
                        { ResultDefList += i + ";"; }
                    }
                    else if (has(Document.Pages[PageIndex].Comment, DefList) == true) // иначе если класс есть в списке ОД
                    {
                        foreach (string i in IsReference(Document.Pages[PageIndex].Comment, DefList)) // добавляем все ОД с совпадениями в список для наложения
                        { ResultDefList += i + ";"; }
                    }
                    if (ResultDefList.Length > 1)
                    { Matching.DefinitionsList = ResultDefList; }
                    Matching.NeedRecognition = true;
                    //Matching.ForceMatch = true; // НЕТ насилию! 
                }
                else // если класса нет в списке ОД
                {
                    if (Matching.DefinitionsList.Contains(@"Template document\Template document")) // если есть универсальный документ
                    {
                        Matching.DefinitionsList = @"Template document\Template document"; // накладываем его!
                        Matching.ForceMatch = true;
                    }
                    // если ничего нет то оставляем как есть, будут накладываться все ОД подряд
                }
            }
            else //если класс не определился
            {
                Matching.DefinitionsList = Matching.DefinitionsList.Replace(@"Template document\Template document;", ""); // убираем из списка ОД универсальный шаблон
                Matching.DefinitionsList = Matching.DefinitionsList.Replace(@"Template document\Template document", "");
                Matching.NeedRecognition = true; 
            }
        }
    }

    public static partial class RoutingRule
    {
        /// <summary>
        /// Допускает на этап, если есть ошибки правил, сборки или количество неуверенных символов >0
        /// выкидывает файлик Verification
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Result"></param>
        public static void NeedVerification(IDocument Document, IRoutingRuleResult Result)
        {

            if (Document.HasErrors || Document.Property("UncertainSymbolsCount") > 0 || Document.AssemblingErrors.Count > 0 || Document.DefinitionName == "" || Document.DefinitionName == "3_42_1_Расшифровка кредиторской задолженности" || Document.DefinitionName == "3_42_2_Расшифровка дебиторской задолженности" || Document.DefinitionName == "3_70_Расшифровка прочих доходов расходов")
            {
                Result.CheckSucceeded = true;
                Utils.ExportXML.Puttxtstatus(Document, "Verification");
            }
            else Result.CheckSucceeded = false;

        }
    }

    public static partial class VerificationTools
    {
        public static void TaskOnClose(ITaskWindow TaskWindow, IBoolean CanClose)
        {
            if (!TaskWindow.Batch.Properties.Has("IgnoreRuleErrors")) //Если нажата кнопка игнорировать ошибки (в меню Сервис окна верификации)
            {
                if (TaskWindow.DocumentsWindow.Items.Count > 0)
                {
                    for (int Docs = 0; Docs < TaskWindow.DocumentsWindow.Items.Count; Docs++)
                    {
                        if (TaskWindow.DocumentState(TaskWindow.DocumentsWindow.Items[Docs].Document).ToString() == "DS_Closed") //если документ закрыт - количество ошибок не доступно
                            TaskWindow.OpenDocument(TaskWindow.DocumentsWindow.Items[Docs].Document); // открываем документ

                        TaskWindow.DocumentsWindow.Items[Docs].Document.CheckRules(); //если ошибки появились после редакции пользователем - система их увидит только после перепроверки правил
                        if (TaskWindow.DocumentsWindow.Items[Docs].Document.DefinitionName.Length > 0)
                        {
                            if (TaskWindow.DocumentsWindow.Items[Docs].Document.HasErrors)
                            {
                                for (int i = 0; i < TaskWindow.DocumentsWindow.Items[Docs].Document.RuleErrors.Count; i++) //проверяем все ошибки - не являются ли они предупреждениями
                                {
                                    if (!TaskWindow.DocumentsWindow.Items[Docs].Document.RuleErrors[i].IsWarning)
                                    {
                                       //FCTools.ShowMessage("В задании остались неисправленные ошибки!");
                                        CanClose.Value = false;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }





        }
    }
}
