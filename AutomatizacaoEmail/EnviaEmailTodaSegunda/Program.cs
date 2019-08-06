using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;

namespace EnviaEmailTodaSegunda
{
    static class Program
    {
        /// <summary>
        /// Ponto de entrada principal para o aplicativo.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            catch(Exception x)
            {
                LogHelper.WriteDebugLog("ERRO:" + x);
            }
            LogHelper.WriteDebugLog("Aplicação encerrada.");
        }

    }
}
