using System;
using System.Windows.Forms;
using Dev4LogLayer.Globals;
using Dev4LogBob50ImportStandardExcelAchVen.Controller;

namespace Dev4LogBob50ImportStandardExcelAchVen
{
    static class Program
    {
        /// <summary> 
        /// THE MAIN ENTRY POINT FOR THE APPLICATION.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            String[] appParams;
            ControllerPrincipal controleurPrincipal = new ControllerPrincipal();

            try
            {
                // LECTURE DES PARAMETRES
                appParams = args[0].Trim().Split(';');
                VarGlobal.runMode = appParams[0];
                VarGlobal.BOBDossierRun = appParams[1];
                VarGlobal.BOBUser = appParams[2];

                // LANCEMENT DE L'APPLICATION
                AppSettingsManager.Load();
                controleurPrincipal.StartApplication();
            }
            catch (Exception e)
            {
                Log.WriteLog("main", "Lancement application (Main).\n" + e.Message);
                MessageBox.Show(e.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
