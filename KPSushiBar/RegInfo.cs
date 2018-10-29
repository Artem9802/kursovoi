using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Data.SqlClient;



namespace KPSushiBar
{
    class RegInfo
    {
       
        public static string ds;
        public static string log;
        public static string pas;
        public SqlConnection Connection = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199");

        public void Register_get()
        {
            try
            {
                RegistryKey Sale_Option = Registry.CurrentConfig;
                RegistryKey DBCon = Sale_Option.CreateSubKey("forSuhi");
              
                Set_Connection();
            }
            catch
            {
                RegistryKey Sale_Option = Registry.CurrentConfig;
                RegistryKey DBCon = Sale_Option.CreateSubKey("forSuhi");
                DBCon.SetValue("ds", "Empty");
                DBCon.SetValue("log", "Empty");
                DBCon.SetValue("pas", "Empty");
            }

        }
        public void Register_set(string DSvalue, string ICvalue, string UNvalue, string UPvalue)
        {
            RegistryKey Sale_Option = Registry.CurrentConfig;
            RegistryKey DBCon = Sale_Option.CreateSubKey("forSuhi");
            Register_get();
            
        }




            public void Set_Connection()
        {

            SqlConnection connection = new SqlConnection("Data Source =" + ds + ";Initial Catalog = master; Persist Security Info = True; User ID = " + log +
            ";Password = \"" + pas + "\"");
        }

        public void Set_Autoriz()
        {
            SqlConnection sc = new SqlConnection("Log_sotr" + log + "Pass_sotr" + pas + "");
        }
    }
}
