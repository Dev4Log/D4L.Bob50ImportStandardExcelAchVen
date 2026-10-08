using System;
using System.Collections.Generic;
using Dev4LogLayer.Bob50.Models;

namespace D4L.Bob50ImportStandardExcelAchVen
{
    public static class VarGlobal
    {
        public static Bob50Folder currentBob50Folder;

        public static string runMode;
        public static string BOBUser;
        public static string BOBLinkPath;
        public static string BOBDossierRun;
        public static string BOBDossierInit;
        public static string DossierBob;
        
        public static List<string> listPdf = new List<string>();
    }
}
