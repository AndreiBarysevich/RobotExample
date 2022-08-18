﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ABBYY.FlexiCapture;
using ABBYY;

namespace SBK
{
    public class NalRes
    {
        #region выбор причины решения нологового органа
        public static string[] reasons1 = new string[]
         {
   $"исполнение обязанности по уплате сумм налогов, сборов, страховых взносов, пеней, штрафов, \r\n указанных в решении о взыскании налогов, сборов, страховых взносов, пеней, штрафов,\r\n процентов за счет денежных средств, а так же электронных денежных средств;",
   $"представление налогоплательщиком (плательщиком сбора, плательщиком страховых взносов,\r\n налоговым агентом) налоговой декларации;",
   $"исполнение обязанности по передаче налоговому органу квитанции о приеме требования \r\nо представлении документов, требования о представлении пояснений и (или) уведомления о вызове в налоговый орган;",
   $"исполнение налогоплательщиком-организацией установленной пунктом 5.1 статьи 23\r\n Налогового кодекса Российской Федерации обязанности по обеспечению получения от налогового органа по месту нахождения организации\r\n (по месту учета организации в качестве крупнейшего налогоплательщика) документов в электронной форме по телекоммуникационным\r\n каналам связи через оператора электронного документооборота",
   $"представление налоговым агентом расчета сумм налога на доходы физических лиц, исчисленных и удержанных налоговым агентом;",
   $"наличие у налогоплательщика (плательщика сбора, плательщика страховых взносов, налогового агента)\r\n денежных средств на счетах в банках, а также электронных денежных средств в размере, превышающем сумму, указанную в решении\r\n о приостановлении операций по счетам налогоплательщика (плательщика сбора, плательщика страховых взносов, налогового агента)\r\n и достаточном для исполнения решения от ________ № ________ о взыскании налогов, сборов, страховых взносов,\r\n пеней, штрафов, процентов за счет денежных средств на счетах в банках, а также электронных денежных средств;",
   $"отмену (замену) обеспечительных мер по основаниям, предусмотренным пунктами 10 и 11\r\n статьи 101 Налогового кодекса Российской Федерации;"
         };

        public static string[] reasons2 = new string[]
        {
$"неисполнением требования об уплате налога, сбора,страховых взносов, пени, штрафа, процентов;",
$"непредставлением налоговой декларации в налоговый орган в течение десяти \r\nрабочих дней по истечении установленного срока ее представления;",
$"обеспечением возможности исполнения решения о привлечении к ответственности \r\nза совершение налогового правонарушения или решения об отказе в привлечении к ответственности за совершение налогового правонарушения;",
$"неисполнением налогоплательщиком-организацией установленной пунктом 5.1 статьи 23\r\n Налогового кодекса Российской Федерации обязанности по передаче налоговому органу квитанции\r\n о приеме требования о представлении документов, требования о представлении пояснений\r\n и (или) уведомления о вызове в налоговый орган - в течение десяти рабочих дней со дня истечения срока, установленного\r\n  для передачи налогоплательщиком-организацией квитанции о приеме документов, направленных налоговым органом;",
$"неисполнением налогоплательщиком-организацией установленной пунктом 5.1 статьи 23\r\n Налогового кодекса Российской Федерации обязанности по обеспечению получения от налогового органа по месту\r\n нахождения организации (по месту учета организации в качестве крупнейшего налогоплательщика) документов\r\n в электронной форме по телекоммуникационным каналам связи через оператора электронного документооборота\r\n - в течение десяти рабочих дней со дня установления налоговым органом факта неисполнения налогоплательщиком\r\n - организацией такой обязанности;",
$"непредставлением налоговым агентом расчета сумм налога на доходы физических лиц,\r\n исчисленных и удержанных налоговым агентом."
        };

        /// <summary>
        /// Метод выдает выпадающий список с значениями причины решения в зависимости от распознаной галки (и проставляет значение по умолчанию)
        /// </summary>
        /// <param name="context"></param>
        public static void DecisionReason(IRuleContext context)
        {
            // Имя полей - меток
            string[] blocks1 = new string[] { "Block", "Block1", "Block2", "Block3", "Block4", "Block5", "Block6", "Block7" };
            string[] blocks2 = new string[] { "Block_1", "Block1_1", "Block2_1", "Block3_1", "Block4_1", "Block5_1" };

            string value1 = null;
            foreach (string i in blocks1)//если хотябы одна галка стоит - запоминаем ее
            {
                if (context.Field(i).Value == true)
                    value1 = i;
                //else value1 = null;
            }

            string value2 = null;
            foreach (string j in blocks2)//если хотябы одна галка стоит - запоминаем ее
            {
                if (context.Field(j).Value == true)
                    value2 = j;
                //else value2 = null;
            }

            if (value1 != null)
            {
                foreach (string i in reasons1)
                    context.Field("DecisionReason").Suggest(i);
                if (context.Field("DecisionReason").IsVerified == false)
                {
                    context.Field("DecisionReason").Value = reasons1[Array.IndexOf(blocks1, value1)];
                    Utils.Region.FocusOn(context, value1, "DecisionReason");
                }
            }
            else if (value2 != null)
            {
                foreach (string i in reasons2)
                    context.Field("DecisionReason").Suggest(i);
                if (context.Field("DecisionReason").IsVerified == false)
                {
                    context.Field("DecisionReason").Value = reasons2[Array.IndexOf(blocks2, value2)];
                    Utils.Region.FocusOn(context, value2, "DecisionReason");
                }
            }
            else
            {
                foreach (string n in reasons1)
                    context.Field("DecisionReason").Suggest(n);
                foreach (string m in reasons2)
                    context.Field("DecisionReason").Suggest(m);
                context.Field("DecisionReason").Value = "";
                Utils.Region.FocusOn(context, "Block", "DecisionReason");
            }
        }
        #endregion
    }
}
