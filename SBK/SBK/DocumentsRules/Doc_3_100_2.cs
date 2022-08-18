using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ABBYY;
using ABBYY.FlexiCapture;

namespace DocumentsRules
{
    public class Doc_3_100_2
    {
        public static void IsBank(IRuleContext Context)
        {
            List<string> notbankwords = new List<string>() { "Покупка иностранной", "Выплата ", "Штрафы", "Налоги и", "Поступления от", "прочие перечисления", "Перевод денежных", "аккредитив", "списание валюты", "банковские доходы", "расчеты с ", "прочие поступления", "прочие расходы", "Услуги банка", "Перевод денежных", "списание валюты", "аренда", "поступление денежных", "Денежные средства в пути", "выручка от", "е услуги", "предпродажная подготовка", "перевозки", "сертификация", "госпошлина", "дивиденды", "курсовая разница", "оплачено за", "расчеты с", "eur", "usd", "Валюта", "Валюта EUR", "Валюта USD", "Валютная сумма", "Внутренние выплаты", "Возврат", "Возврат краткосрочных кредитов", "инкассация", "Итого", "Налог на имущество", "Оплата", "Оплата Услуг", "Оплата штрафов, пени", "Получение займа", "Поступление", "Поступление платежей", "Выдача под", "Выплата", "услуги", "Отчисления", "Платежи", "Покупка", "наличных", "Получение", "Проценты", "Прочие", "Выдача" };
            if (Context.HasField("Subkonto") && Context.Field("Schet").IsVerified == false && !string.IsNullOrEmpty(Context.Field("Subkonto").Text))
            {
                foreach (string word in notbankwords)
                {
                    if (Context.Field("Subkonto").Text.ToLower().Replace(@"\r\n", "").Contains(word.ToLower())
                    ) //ищем строки не содержащие слова из списка
                    {
                        Context.Field("Schet").Value = false;

                        break;
                    }
                    else
                    {
                        Context.Field("Schet").Value = true;
                        //Context.Field("Subkonto").IsVerified = true;
                    }

                }
            }
        }



    }
    public class Doc_3_55
    {
        public static void IsSchet(IRuleContext Context)
        {
            List<string> notschetkwords = new List<string>() {"Оплата", "выплаты", "внутренние выплаты", "50", "51", "поступление", "реализаци", "услуг" };
            if (Context.HasField("Acc") && Context.Field("Schet").IsVerified == false && !string.IsNullOrEmpty(Context.Field("Acc").Text))
            {
                foreach (string word in notschetkwords)
                {
                    if (Context.Field("Acc").Text.ToLower().Replace(@"\r\n", "").Contains(word.ToLower())
                    ) //ищем строки не содержащие слова из списка
                    {
                        Context.Field("Schet").Value = false;
                        break;
                    }
                    else 
                    {
                        Context.Field("Schet").Value = true;
                        //Context.Field("Subkonto").IsVerified = true;
                    }

                }
            }
        }

    }
}
