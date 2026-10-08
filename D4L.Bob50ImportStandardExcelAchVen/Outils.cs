using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace D4L.Bob50ImportStandardExcelAchVen
{
    static class Outils
    {

        public static string Left(string StrTexte, int LenTexte)
        {

            string cTemp = null;

            if (string.IsNullOrEmpty(StrTexte))
                cTemp = "";
            else
                cTemp = StrTexte.Substring(0, LenTexte);
            return (cTemp);
        }

        public static string Right(string StrTexte, int LenTexte)
        {

            string cTemp = null;

            if (string.IsNullOrEmpty(StrTexte))
                cTemp = "";
            else
            {
                if (StrTexte.Length <= LenTexte)
                    cTemp = StrTexte;
                else
                    cTemp = StrTexte.Substring(StrTexte.Length - LenTexte);
            }
            return (cTemp);
        }


        public static int ExecCmd(string CmdeDir, string CmdeLine, string CmdeParam)
        {

            Process ExecProcess = new Process();

            // Get the path that stores user documents.

            try
            {
                ExecProcess.StartInfo.WorkingDirectory = CmdeDir;
                ExecProcess.StartInfo.FileName = CmdeLine;
                ExecProcess.StartInfo.Arguments = CmdeParam;
                ExecProcess.StartInfo.UseShellExecute = true;
                ExecProcess.Start();
                ExecProcess.WaitForExit();
                // This code assumes the process you are starting will terminate itself. 
                // Given that is is started without a window so you cannot terminate it 
                // on the desktop, it must terminate itself or you can do it programmatically
                // from this application using the Kill method.
                return (0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (1);
            }
        }

        public static bool IsDate(string InputDate)
        {
            DateTime TempDT;
            return (DateTime.TryParse(InputDate, out TempDT));
        }

        public static bool IsNumeric(string InputNum, string TypeNum)
        {
            int TempInt;
            double TempDouble;

            try
            {
                if (TypeNum == "INT")
                    TempInt = Convert.ToInt32(InputNum);
                else
                    TempDouble = Convert.ToDouble(InputNum);
                return (true);
            }
            catch (Exception)
            {
                return (false);
            }
        }

        public static Boolean checkIban(string iban)
        {
            string ibanTemp;
            int longueurString;
            int compteur;
            char ch; ;
            Int32 checksum = 0;

            ibanTemp = iban.ToUpper();
            ibanTemp = iban.Replace(" ", String.Empty);
            if ((ibanTemp.Length < 15) || (ibanTemp.Length > 34))
                return false;

            // le code pays retourne à l’arrière
            ibanTemp = ibanTemp.Substring(4, ibanTemp.Length - 4) + ibanTemp.Substring(0, 4);

            // Conversion des valeurs de lettres selon convention A=10, etc. 
            compteur = 0;
            do
            {
                ch = ibanTemp[compteur];
                if ((int)ch > 64 && (int)ch < 91)
                {
                    ibanTemp = ibanTemp.Replace(ch.ToString(), ((int)ch - 55).ToString());
                }
                compteur++;
                longueurString = ibanTemp.Length;
            }
            while (compteur < longueurString);

            // Check sur valeur de modulo 97
            for (int i = 0; i <= ibanTemp.Length - 1; i++)
                checksum = (checksum * 10 + Convert.ToInt32(ibanTemp.Substring(i, 1))) % 97;

            if (checksum == 1)
                return true;
            else
                return false;
        }
    }
}
