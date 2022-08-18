using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Config
{
    public static class Values
    {
        /// <summary>
        /// Строка подключения к БД справочников.
        /// </summary>
        public static string DictionaryConnString =
            @"Data Source=localhost\SQLEXPRESS;Initial Catalog=BD;Integrated Security=False;User ID=abbyyuser;Password=q1w2e3r4;";

        /// <summary>
        /// Строка подключения к БД FlexiCapture.
        /// </summary>
        public static string FcConnectionString =
            @"Data Source=localhost\SQLEXPRESS;Initial Catalog=AbbyyDB;Integrated Security=False;User ID=abbyyuser;Password=q1w2e3r4;";


    }
}
