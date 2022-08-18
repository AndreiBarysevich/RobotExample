using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ABBYY.FlexiCapture;
using ABBYY;
using Utils;

namespace DocumentsRules
{
    /// <summary>
    /// Отчет о целевом использовании средств
    /// </summary>
    public class Doc_3_341
    {
        /// <summary>
        /// Метод проверяет наличие обязательных кодов в таблице - не работает
        /// </summary>
        /// <param name="Context"></param>
        //public static void HasRequiredKods(IRuleContext Context)
        //{
        //    bool i6400 = false;
        //    bool i6100 = false;
        //    for (var i = 0; i < Context.Field("Table").Items.Count; i++)
        //    {
        //        if (Context.Field("Kod").Items[i].Text.Contains("6100") || Context.Field("Kod").Items[i].Text.Contains("6100"))
        //        {
        //            if (Context.Field("Kod").Items[i].Text.Contains("6100")) { i6100 = true; continue; }
        //            if (Context.Field("Kod").Items[i].Text.Contains("6400")) { i6400 = true; continue; }

        //        }
        //        else continue;
        //    }

        //    if (i6400 == false || i6100 == false)
        //    {
        //        Context.CheckSucceeded = false;
        //        Context.ErrorMessage = "Таблица должна обязательно содержать Остаток на начало года (6100) и на конец года (6400)";
        //    }

        //}


    }
}
