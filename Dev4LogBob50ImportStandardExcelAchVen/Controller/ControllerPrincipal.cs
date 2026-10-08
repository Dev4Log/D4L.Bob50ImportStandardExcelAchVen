using System;
using System.IO;
using System.Configuration;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using Dev4LogLayer.Bob50;
using Dev4LogLayer.Globals;
using Dev4LogLicence = Dev4LogLayer.License;
using Dev4LogBob50ImportStandardExcelAchVen.DAO;


namespace Dev4LogBob50ImportStandardExcelAchVen.Controller
{
    
    public class ControllerPrincipal
    {

        Bob50Dao dev4LogBob50Dao = new Bob50Dao();

        public ControllerPrincipal()
        {
        }

        /// <summary>
        /// CONTROLES PREALABLE ET LANCEMENT DE L'APPLICATION
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void StartApplication()
        {
            string licenceMsg = "";
            bool licenceOk = true;

            #region VALIDATION DE LA LICENCE

            DateTime licenceDate = Dev4LogLicence.ControllerLicense.CheckLicense();
            if (licenceDate < DateTime.Now)
            {
                licenceOk = false;
                if (licenceDate == new DateTime(2, 1, 1))
                    licenceMsg = "Licence not found";
                else
                {
                    if (licenceDate == new DateTime(1, 1, 1))
                    {
                        licenceMsg = "Invalid Licence";
                    }
                    else
                    {
                        if (licenceDate > DateTime.Today)
                        {
                            licenceMsg = $@"Expired Licence ({licenceDate.ToString("yyyy-MM-dd")}), please renew it from the licence menu";
                        }
                    }
                }

                if (VarGlobal.runMode != "AUTO")
                {
                    if (!string.IsNullOrEmpty(licenceMsg))
                        MessageBox.Show(licenceMsg, "Dev4Log", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Log.WriteLog("Licence", licenceMsg);
                    return;
                }
            }

            licenceOk = true;

            #endregion

            #region VALIDATION DU REPERTOIRE BOB50

            if (!File.Exists(Path.Combine(AppSettingsManager.Dictionnary["BOB_PATH"], "BOB.EXE")))
            {
                throw new Exception("Répertoire Bob50 invalide (" + AppSettingsManager.Dictionnary["BOB_PATH"] + ")");
            }

            #endregion

            #region LECTURE BOB.INI

            ReadBobIni.Bob50Directory = AppSettingsManager.Dictionnary["BOB_PATH"];
            ReadBobIni.Read();
            VarGlobal.BOBLinkPath = ReadBobIni.LinkDirectory;

            #endregion

            #region LANCEMENT DE L'APPLICATION

            VarGlobal.currentBob50Folder = dev4LogBob50Dao.GetBobFolder(VarGlobal.BOBDossierRun);
            Controller.ControllerImportExcel CtrlImportsExcel = new Controller.ControllerImportExcel(licenceOk);

            #endregion
        }
    }
}
