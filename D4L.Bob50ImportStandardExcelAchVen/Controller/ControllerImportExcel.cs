using System;
using System.Configuration;
using Dev4LogLayer.Bob50;
using Dev4LogLayer.Bob50.Models;
using Dev4LogLayer.Globals;
using Dev4LogLayerFiles = Dev4LogLayer.Files;
using D4L.Bob50ImportStandardExcelAchVen.GUI;
using System.Windows.Forms;
using System.IO;
using D4L.Bob50ImportStandardExcelAchVen.Models;
using System.Collections.Generic;
using System.Data;
using System.Collections;
using D4L.Bob50ImportStandardExcelAchVen.DAO;
using static System.Net.Mime.MediaTypeNames;
using System.Net;
using System.Text.RegularExpressions;
using Dev4LogLayer.Globals.GUI;
using System.Text;

namespace D4L.Bob50ImportStandardExcelAchVen.Controller
{
    public class ControllerImportExcel
    {
        Bob50Dao dev4LogBob50dao = new Bob50Dao();
        private readonly ImportsExcel currForm;

        /// <summary>
        /// // INITIALISATION DE LA FORME
        /// </summary>        
        public ControllerImportExcel(bool licenceOk)
        {
            try
            {
                currForm = new ImportsExcel(this);

                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                currForm.Text = $"{currForm.Text} ({version.Major}.{version.Minor}.{version.Build}.{version.Revision}) (01/2026)";

                if (VarGlobal.runMode == "MULTI")
                {
                    currForm.GroupBox1.Text = " ** IMPORT MULTI-DOSSIERS ** ";
                    currForm.ButtExecuter.Visible = false;
                    currForm.ButtExecuterMulti.Left = currForm.ButtExecuter.Left;
                    currForm.ButtExecuterMulti.Top = currForm.ButtExecuter.Top;
                    currForm.ButtExecuterMulti.Visible = true;
                    if (!licenceOk)
                    {
                        currForm.ButtExecuterMulti.Enabled = false;
                    }
                }
                else
                {
                    currForm.toolStripStatusLabel1.Text = dev4LogBob50dao.GetBobFolder(VarGlobal.BOBDossierRun).Name + " [" + VarGlobal.BOBDossierRun + "]";
                    if (!licenceOk)
                    {
                        currForm.ButtExecuter.Enabled = false;
                    }
                }

                // AFFICHAGE DE LA FORME
                currForm.ShowDialog();
            }
            catch (Exception e)
            {
                Log.WriteLog("ControllerImportExcel", "Constructeur() : " + e.Message);
            }
        }


        /// <summary>
        /// OUVERTURE FENETRE DE DIALOGUE DE RECHERCHE DU FICHIER EXCEL A IMPORTER
        /// </summary>
        public void SelectFichierExcel()
        {
            OpenFileDialog SelectXlsFileDialog = new OpenFileDialog();
            SelectXlsFileDialog.InitialDirectory = AppSettingsManager.Dictionnary["IMPORT_PATH"];
            SelectXlsFileDialog.Filter = "Excel files |*.xlsx";
            SelectXlsFileDialog.RestoreDirectory = true;
            SelectXlsFileDialog.ShowDialog();
            currForm.FichierImport.Text = SelectXlsFileDialog.FileName;
        }


        /// <summary>
        /// IMPORT EXCEL STANDARD
        /// </summary>
        public void RunImportStandard()
        {
            const int BIE_JOURNAL = 0;
            const int BIE_ANNEE = 1;
            const int BIE_MOIS = 2;
            const int BIE_DATEDOC = 3;
            const int BIE_DATEECH = 4;
            const int BIE_NUMDOC = 5;
            const int BIE_TIERS = 6;
            const int BIE_COMPTE = 7;
            const int BIE_REM = 8;
            const int BIE_COMM = 9;
            const int BIE_TOTALTTC = 10;
            const int BIE_NATTVA = 11;
            const int BIE_PRCTVA = 12;
            const int BIE_BASE = 13;
            const int BIE_TVA = 14;
            const int BIE_DEVISE = 15;
            const int BIE_COURS = 16;
            const int BIE_TIERSNOM = 17;
            const int BIE_TIERSADR = 18;
            const int BIE_TIERSCP = 19;
            const int BIE_TIERSVILLE = 20;
            const int BIE_TIERSASS = 21;
            const int BIE_TIERSTVA = 22;
            const int BIE_TIERSEMAIL = 23;
            const int BIE_NAMEDOC = 24;
            const int BIE_MANDAT = 25;
            const int BIE_TIERSIBAN = 26;
            const int BIE_TIERSBIC = 27;
            const int BIE_TIERSBQDEF = 28;

            int NbColExcel;
            int NbColStd;

            TVAInfos tvainfos = null;

            List<string> ListeSecAna = new List<string>();
            List<string> ListeTypeSecAna = new List<string>();
            List<Boolean> ListeSecAnaMandatory = new List<Boolean>();

            Bob50KhDbk HeaderACH = new Bob50KhDbk();
            Bob50KhDbk HeaderNCA = new Bob50KhDbk();
            Bob50KhDbk HeaderVEN = new Bob50KhDbk();
            Bob50KhDbk HeaderNCV = new Bob50KhDbk();

            Bob50KlDbk LineACH = new Bob50KlDbk();
            Bob50KlDbk LineNCA = new Bob50KlDbk();
            Bob50KlDbk LineVEN = new Bob50KlDbk();
            Bob50KlDbk LineNCV = new Bob50KlDbk();

            int NumLigne;

            string RuptureDoc = "";
            int IndImp = 0;
            double BaseTva = 0;
            double Tva = 0;
            double TvaDeduc = 0;
            double TotalTTC = 0;
            double TotalBase = 0;
            double TotalTva = 0;
            double Ecart = 0;
            bool FichierOk;
            bool LigneOk;
            string VatType = "";
            string VatNat = "";
            double VatPrc = 0;
            double CoursDevise = 0;
            DateTime DateDoc;
            DateTime DateEch;

            List<DmDoc> listDmDoc = new List<DmDoc>();

            bool iBanExist;
            Int32 iBanOrder;

            bool LinkOk;
            var RetVal = 0;
            string cTemp = "";

            Boolean AnaObligatoireGlobale;
            Boolean TvaActive;

            // ARRAY POUR LECTURE LIGNE EXCELL
            List<string> ExcellArray = new List<string>();

            try
            {

                // TEST SI FICHIER D'IMPORT BIEN SELECTIONNE
                if (String.IsNullOrEmpty(currForm.FichierImport.Text))
                {
                    MessageBox.Show("Veuillez d'abord sélectionner un fichier Excel à importer", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (!File.Exists(currForm.FichierImport.Text))
                {
                    MessageBox.Show("Fichier d'import inexistant", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // CONFIRMATION DE LA LIAISON
                if (MessageBox.Show("Démarrer l'import", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
                    return;

                // CREATION DU REPERTOIRE LINK DU DOSSIER SI NECESSAIRE ET COPIE DES FICHIERS DE LIAISON
                Bob50Dao bob50dao = new Bob50Dao();
                Bob50Folder folder = bob50dao.GetBobFolder(VarGlobal.BOBDossierRun);
                //Bob50Config conf = bob50dao.GetConfigDossier(folder);
                Bob50Config conf = DaoBob.LOCAL_GetConfigDossier(folder);

                // LECTURE DU FICHIER EXCEL
                Dev4LogLayer.Files.Models.Excel xlsx = new()
                {
                    FilePath = currForm.FichierImport.Text,
                    OnlyString = false
                };
                DataTable tableFichierExcel = Dev4LogLayerFiles.Excel.ExcelToDataTable(xlsx);

                currForm.ProgressBar1.Value = 01;
                currForm.ProgressBar1.Visible = true;

                // INITIALISATIONS
                int NumDocACH = 990000;
                int NumDocNCA = 990000;

                int NumDocVEN = 0;
                int NumDocNCV = 0;

                // PARCOURS DU FICHIER EXCEL
                currForm.InfosImport.Text = "";
                RuptureDoc = "";
                FichierOk = true;

                #region GESTION DE LA PARTIE ANALYTIQUE

                // INITIALISATION NOMBRE DE COLONNES ET DES CODES DES SECTIONS ANALYTIQUES
                NbColExcel = tableFichierExcel.Columns.Count;
                if (VarGlobal.runMode == "DOM")
                    NbColStd = 28;
                else
                    NbColStd = 24;
                for (int i = NbColStd + 1; i < tableFichierExcel.Columns.Count; i++)
                {
                    ListeSecAna.Add(tableFichierExcel.Columns[i].ColumnName);
                }

                for (int i = 0; i < ListeSecAna.Count; i++)
                {
                    string typeSection = DaoBob.GetTypeSecAna(ListeSecAna[i], folder.Path);
                    string sectionObligatoire = DaoBob.GetSecAnaObligatoire(ListeSecAna[i], folder.Path);
                    switch (sectionObligatoire)
                    {
                        case "TRUE":
                            ListeSecAnaMandatory.Add(true);
                            ListeTypeSecAna.Add(typeSection);
                            break;

                        case "FALSE":
                            ListeSecAnaMandatory.Add(false);
                            ListeTypeSecAna.Add(typeSection);
                            break;
                        default:
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "SECTION ANALYTIQUE INEXISTANTE : " + ListeSecAna[i] + (char)(13) + (char)(10);
                            FichierOk = false;
                            break;
                    }
                    if (FichierOk == false)
                        break;
                }

                // CREATION DES FICHIERS DE LIAISON
                bob50dao.CreateBobLinkFileEntry(folder, "ACH");
                bob50dao.CreateBobLinkFileEntry(folder, "NCV");
                bob50dao.CreateBobLinkFileEntry(folder, "NCA");
                bob50dao.CreateBobLinkFileEntry(folder, "VEN");

                // SORTIE APPLICATION SI PROBLEME DE CONFIGURATION ANALYTIQUE
                if (!FichierOk)
                {
                    MessageBox.Show("Erreurs dans la configuation analytique du fichier d'import, veuillez corriger", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                #endregion

                #region LECTURE PARAMETRES DOSSIER ENCODAGE AVEC TVA ET ANALYTIQUE OBLIGATOIRE MODE GLOBAL

                String DeviseBaseDossier = conf.Basecurrid;
                String LegisDossier = conf.Legis;
                AnaObligatoireGlobale = conf.Costmand;
                TvaActive = conf.Prmvatenbld;

                #endregion

                #region LECTURE ET CONTROLE DE COHERENCE DU FICHIER EXCEL

                NumLigne = 2;
                foreach (DataRow dr in tableFichierExcel.Rows)
                {
                    // LECTURE LIGNE EXCELL
                    ExcellArray.Clear();
                    for (int i = 0; i < NbColExcel; i++)
                    {
                        if (!DBNull.Value.Equals(dr[i]))
                            ExcellArray.Add(dr[i].ToString().Trim());
                        else
                            ExcellArray.Add("");
                    }

                    // CONTROLES ET AJOUT/MODIFICATIONS PREALABLES
                    // -------------------------------------------
                    LigneOk = true;

                    Bob50Dbk journalInfos = new Bob50Dbk();

                    // JOURNAL
                    if (ExcellArray[BIE_JOURNAL] == "")
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> JOURNAL VIDE" + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }
                    if (!DaoBob.GetJournalOk(ExcellArray[BIE_JOURNAL], folder.Path))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> JOURNAL INEXISTANT : " + ExcellArray[BIE_JOURNAL] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                        continue;
                    }
                    else
                    {
                        journalInfos = DaoBob.GetJournalInfos(folder, ExcellArray[BIE_JOURNAL], folder.Path);
                    }

                    if (ExcellArray[BIE_NAMEDOC] != "" && !File.Exists(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"] + ExcellArray[BIE_NAMEDOC]))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> Le document pdf entré n'existe pas." + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // PERIODE
                    Bob50Period periode = DaoBob.GetPeriode(ExcellArray[BIE_ANNEE], ExcellArray[BIE_MOIS], folder.Path);
                    if (periode.Month == 0 || periode.Year == 0)
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> PERIODE INEXISTANTE : " + ExcellArray[BIE_MOIS] + " " + ExcellArray[BIE_ANNEE] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // DATE
                    if (!Outils.IsDate(ExcellArray[BIE_DATEDOC]))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE INVALIDE : " + ExcellArray[BIE_DATEDOC] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                        DateDoc = DateTime.Now;
                    }
                    else
                    {
                        DateDoc = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                        if ((journalInfos.Dbtype == "SAL") || (journalInfos.Dbtype == "SAC"))
                        {
                            if (DateDoc.Month != periode.Month || DateDoc.Year != periode.Year)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE N'APPARTIENT PAS A LA PERIODE : " + ExcellArray[BIE_DATEDOC] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }
                        else
                        {
                            if (DateDoc.Month > periode.Month && DateDoc.Year == periode.Year)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE SUPERIEURE A LA PERIODE : " + ExcellArray[BIE_DATEDOC] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }
                    }

                    // DATE ECHEANCE
                    if (ExcellArray[BIE_DATEECH] == "")
                        DateEch = DateDoc;
                    else
                    {
                        if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                        {
                            DateEch = Convert.ToDateTime(ExcellArray[BIE_DATEECH]);
                            if (DateEch < DateDoc)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE D'ECHEANCE INFERIEURE A LA DATE DU DOCUMENT : " + ExcellArray[BIE_DATEECH] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }
                        else
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE D'ECHEANCE INVALIDE : " + ExcellArray[BIE_DATEECH] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                            DateEch = DateTime.Now;
                        }
                    }

                    // AJOUT MODIFICATION TIERS SI NECESSAIRE, PARTIE SIGNALETIQUE ET TVA
                    if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSNOM]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSADR]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSCP]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSVILLE]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSASS]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSTVA]))
                    {
                        Tuple<bool, bool> resultTva = DaoBob.VerifStdCode("VATCOUNT", ExcellArray[BIE_TIERSASS], ReadBobIni.CommonDirectory);
                        Bob50Tiers ajoutCompan = new Bob50Tiers();
                        if (ExcellArray[BIE_TIERSASS] == "EX" || ExcellArray[BIE_TIERSASS] == "NA" || resultTva.Item1)
                        {
                            ajoutCompan.Cid = ExcellArray[BIE_TIERS];
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSNOM]))
                                ajoutCompan.Cname1 = ExcellArray[BIE_TIERSNOM].Replace("'", "''");
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSADR]))
                                ajoutCompan.Cadress1 = ExcellArray[BIE_TIERSADR].Replace("'", "''");
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSCP]))
                                ajoutCompan.Czipcode = ExcellArray[BIE_TIERSCP];
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSVILLE]))
                                ajoutCompan.Clocality = ExcellArray[BIE_TIERSVILLE];
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSEMAIL]))
                                ajoutCompan.Emailaddress = ExcellArray[BIE_TIERSEMAIL];
                            switch (ExcellArray[BIE_TIERSASS])
                            {
                                case "NA":
                                    ajoutCompan.Cvatcat = "N";
                                    ajoutCompan.Cvatref = "";
                                    ajoutCompan.Cvatno = "";
                                    break;
                                case "EX":
                                    ajoutCompan.Cvatcat = "";
                                    ajoutCompan.Cvatref = "EX";
                                    ajoutCompan.Cvatno = ExcellArray[BIE_TIERSTVA];
                                    break;
                                default:
                                    ajoutCompan.Cvatcat = "";
                                    ajoutCompan.Cvatref = ExcellArray[BIE_TIERSASS];
                                    ajoutCompan.Cvatno = ExcellArray[BIE_TIERSTVA];
                                    break;
                            }

                            if (!DaoBob.IfExistTier(ExcellArray[BIE_TIERS], folder.Path))
                            {
                                if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                                {
                                    ajoutCompan.Ccustype = "C";
                                    ajoutCompan.Csuptype = "U";
                                }
                                else
                                {
                                    ajoutCompan.Ccustype = "U";
                                    ajoutCompan.Csuptype = "S";
                                }
                                DaoBob.AjoutCompan(ajoutCompan, folder.Path);
                            }
                            else
                            {
                                if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                                {
                                    ajoutCompan.Ccustype = "C";
                                    DaoBob.UpdateCompan(ajoutCompan, "C", folder.Path);
                                }
                                else
                                {
                                    ajoutCompan.Csuptype = "S";
                                    DaoBob.UpdateCompan(ajoutCompan, "S", folder.Path);
                                }
                            }

                            if (AppSettingsManager.Dictionnary["MODE_IMPORT"] == "DOM")
                            {
                                if (!string.IsNullOrEmpty(ExcellArray[BIE_MANDAT]))
                                {
                                    ajoutCompan.Cbankorderb2b = true;
                                    ajoutCompan.Cbankorderpay = true;
                                    ajoutCompan.Cbankordermandate = ExcellArray[BIE_MANDAT];
                                    DaoBob.UpdateCompanBank(ajoutCompan, folder.Path);
                                }
                            }

                        }
                        else
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TYPE TVA TIERS INCORRECT (NA/EX/BE, LU, ...) : " + ExcellArray[BIE_TIERSASS] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                    }

                    if (AppSettingsManager.Dictionnary["MODE_IMPORT"] == "DOM")
                    {
                        // AJOUT MODIFICATION TIERS PARTIE IBAN
                        if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSIBAN]))
                        {
                            if (Outils.checkIban(ExcellArray[BIE_TIERSIBAN]))
                            {
                                // MISE A JOUR PARTIE TIERS BANQUE PAR DEFAUT
                                if (ExcellArray[BIE_TIERSBQDEF] == "*")
                                {
                                    if (DaoBob.IfExistTier(ExcellArray[BIE_TIERS], folder.Path))
                                    {
                                        if (LegisDossier == ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2))
                                        {
                                            String cbankno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                            String cbankiban = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                            String cbankcode = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                            String cbnktypepay = "NAT";
                                            DaoBob.UpdateTierBankInfos(ExcellArray[BIE_TIERS], cbankno, cbankiban, cbankcode, cbnktypepay, folder.Path);
                                        }
                                        else
                                        {
                                            DaoBob.UpdateTierBankInfos(ExcellArray[BIE_TIERS], "", "", "", "NAT", folder.Path);
                                        }
                                    }
                                }

                                // MISE A JOUR PARTIE BANQUE
                                iBanExist = false;
                                iBanOrder = 1 + DaoBob.IfExistCompBk(ExcellArray[BIE_TIERS], folder.Path);
                                CompBk ajoutCompBk;

                                if (iBanOrder > 1)
                                {
                                    if (DaoBob.IfCompBkIbanExist(ExcellArray[BIE_TIERS], ExcellArray[BIE_TIERSIBAN].Replace(" ", ""), folder.Path))
                                    {
                                        iBanExist = true;
                                        ajoutCompBk = new CompBk();
                                        ajoutCompBk.Cid = ExcellArray[BIE_TIERS];
                                        ajoutCompBk.Banknoindex = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");

                                        if (ExcellArray[BIE_TIERSBQDEF] == "*")
                                            ajoutCompBk.Cdefault = "true";
                                        else
                                            ajoutCompBk.Cdefault = "false";

                                        if (LegisDossier == ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2))
                                        {
                                            ajoutCompBk.Type = "NAT";
                                            ajoutCompBk.Cbankno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                            ajoutCompBk.Cbankcode = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                            ajoutCompBk.Cbankiban = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                        }
                                        else
                                        {
                                            ajoutCompBk.Type = "FOR";
                                            ajoutCompBk.Cforbnkno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                            ajoutCompBk.Cforbnkswift = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                            ajoutCompBk.Cforbnkctry = ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2);
                                        }

                                        DaoBob.UpdateCompBk(ajoutCompBk, folder.Path);
                                    }

                                    if (ExcellArray[BIE_TIERSBQDEF] == "*")
                                    {
                                        DaoBob.UpdateCompBkDefault(ExcellArray[BIE_TIERS], ExcellArray[BIE_TIERSIBAN].Replace(" ", ""), "false", folder.Path);
                                    }

                                }

                                // AJOUT NOUVELLE BANQUE
                                if (iBanExist == false)
                                {
                                    CompBk compbk = new CompBk();
                                    compbk.Cid = ExcellArray[BIE_TIERS];
                                    compbk.Order = iBanOrder.ToString();
                                    compbk.Banknoindex = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");

                                    // MISE A JOUR DES CHAMPS COMPLEMENTAIRES
                                    if (LegisDossier == ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2))
                                    {
                                        compbk.Type = "NAT";
                                        compbk.Cbankno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                        compbk.Cbankiban = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                        compbk.Cbankcode = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                    }
                                    else
                                    {
                                        compbk.Type = "FOR";
                                        compbk.Cforbnkno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                        compbk.Cforbnkswift = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                        compbk.Cforbnkctry = ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2);
                                    }
                                    if (ExcellArray[BIE_TIERSBQDEF] == "*")
                                        compbk.Cdefault = "true";
                                    else
                                        compbk.Cdefault = "false";

                                    DaoBob.AjoutCompBk(compbk, folder.Path);
                                }

                                // MISE A JOUR PARTIE TIERS DE
                                CompDe cmpde = new CompDe();
                                cmpde.Cid = ExcellArray[BIE_TIERS];
                                if (DaoBob.IfExistCompDe(cmpde.Cid, folder.Path))
                                {

                                    // MISE A JOUR DES CHAMPS COMPLEMENTAIRES
                                    if (LegisDossier == ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2))
                                    {
                                        cmpde.Cforbnkno = "";
                                        cmpde.Cforbnkswift = "";
                                        cmpde.Cforbnkctry = "";
                                    }
                                    else
                                    {
                                        cmpde.Cforbnkno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                        cmpde.Cforbnkswift = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                        cmpde.Cforbnkctry = ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2);
                                    }
                                    DaoBob.AjoutUpdateCompDe(cmpde, VarGlobal.BOBDossierRun);
                                }
                                else
                                {
                                    if (LegisDossier != ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2))
                                    {
                                        // MISE A JOUR DES CHAMPS COMPLEMENTAIRES
                                        cmpde.Cforbnkno = ExcellArray[BIE_TIERSIBAN].Replace(" ", "");
                                        cmpde.Cforbnkswift = ExcellArray[BIE_TIERSBIC].Replace(" ", "");
                                        cmpde.Cforbnkctry = ExcellArray[BIE_TIERSIBAN].Replace(" ", "").Substring(0, 2);
                                        DaoBob.AjoutUpdateCompDe(cmpde, folder.Path);
                                    }
                                }
                            }
                            else
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> IBAN INCORRECT : " + ExcellArray[BIE_TIERSIBAN] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }
                    }

                    // TIERS & INITIALISATION TYPE TVA DU TIERS POUR CONTROLE DU CODE TVA
                    VatType = "N";
                    if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                    {
                        Bob50Tiers customer = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "C", folder.Path);
                        if (customer == null)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> CLIENT INEXISTANT : " + ExcellArray[BIE_TIERS] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            if (customer.Cvatcat != "N")
                            {
                                if (customer.Cvatref != LegisDossier)
                                {
                                    cTemp = customer.Cvatref;
                                    switch (cTemp)
                                    {
                                        case "EX":
                                            VatType = "I";
                                            break;
                                        case "":
                                            VatType = "N";
                                            break;
                                        default:
                                            VatType = "E";
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Bob50Tiers supplier = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "S", folder.Path);
                        if (supplier == null)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> FOURNISSEUR INEXISTANT : " + ExcellArray[BIE_TIERS] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            if (supplier.Cvatcat != "N")
                            {
                                if (supplier.Cvatref != LegisDossier)
                                {
                                    cTemp = supplier.Cvatref;
                                    switch (cTemp)
                                    {
                                        case "EX":
                                            VatType = "I";
                                            break;
                                        case "":
                                            VatType = "N";
                                            break;
                                        default:
                                            VatType = "E";
                                            break;
                                    }
                                }
                            }
                        }
                    }

                    // COMPTE GENERAL
                    DataTable accInfos = DaoBob.getAccountInfos(ExcellArray[BIE_COMPTE], folder.Path);
                    if (accInfos == null)
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COMPTE GENERAL INEXISTANT : " + ExcellArray[BIE_COMPTE] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }
                    else
                    {
                        bool aistitle = Convert.ToBoolean(accInfos.Rows[0]["AISTITLE"]);
                        if (aistitle)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COMPTE GENERAL INVALIDE : " + ExcellArray[BIE_COMPTE] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            // CONTROLE DE L'ANALYTIQUE
                            for (int i = 0; i < ListeSecAna.Count; i++)
                            {
                                bool accCost = Convert.ToBoolean(accInfos.Rows[0]["COSTR" + ListeSecAna[i]]);
                                if (accCost)
                                {
                                    if (ExcellArray[NbColStd + 1 + i].Trim() == "")
                                    {
                                        if (ListeSecAnaMandatory[i] || AnaObligatoireGlobale)
                                        {
                                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> ANALYTIQUE : " + ListeSecAna[i] + " OBLIGATOIRE" + (char)(13) + (char)(10);
                                            FichierOk = false;
                                            LigneOk = false;
                                        }
                                    }
                                    else
                                    {
                                        if (ListeTypeSecAna[i] == "")
                                        {
                                            Tuple<bool, bool> cosecInfos = DaoBob.getCostAnaInfos(ListeSecAna[i], ExcellArray[NbColStd + 1 + i], folder.Path);
                                            if (!cosecInfos.Item1)
                                            {
                                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> CODE ANALYTIQUE INVALIDE : " + ExcellArray[NbColStd + 1 + i] + (char)(13) + (char)(10);
                                                FichierOk = false;
                                                LigneOk = false;
                                            }
                                            else
                                            {
                                                if (cosecInfos.Item2)
                                                {
                                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> CODE ANALYTIQUE INVALIDE : " + ExcellArray[NbColStd + 1 + i] + (char)(13) + (char)(10);
                                                    FichierOk = false;
                                                    LigneOk = false;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (ExcellArray[NbColStd + 1 + i].Trim() != "")
                                    {
                                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COMPTE GENERAL NON ANALYTIQUE : " + ExcellArray[BIE_COMPTE] + " [" + ListeSecAna[i] + "]" + (char)(13) + (char)(10);
                                        FichierOk = false;
                                        LigneOk = false;
                                    }
                                }
                            }
                        }
                    }

                    // MONTANT TVAC
                    if (!Outils.IsNumeric(ExcellArray[BIE_TOTALTTC], "DOUBLE"))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> MONTANT TVAC INVALIDE : " + ExcellArray[BIE_TOTALTTC] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // TVA
                    if (TvaActive && journalInfos.Iswithvat)
                    {
                        ExcellArray[BIE_PRCTVA].Replace("%", "");
                        if (!Outils.IsNumeric(ExcellArray[BIE_PRCTVA], "DOUBLE"))
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TAUX TVA INVALIDE : " + ExcellArray[BIE_PRCTVA] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            VatNat = ExcellArray[BIE_NATTVA];
                            VatPrc = Convert.ToDouble(ExcellArray[BIE_PRCTVA]);
                            tvainfos = DaoBob.GetVATInfos("S", VatType, VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                            if (tvainfos == null)
                            {
                                if (VatType == "N" && (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC"))
                                {
                                    // SI CLIENT NA -> CONTROLE TVA CAS VENTES WEB -> TVA AU TAUX DU PAYS DU CLIENT OU EXPORT
                                    tvainfos = DaoBob.GetVATInfos("S", "E", VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                                    if (tvainfos == null)
                                    {
                                        if (tvainfos == null)
                                        {
                                            tvainfos = DaoBob.GetVATInfos("S", "I", VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                                            if (tvainfos == null)
                                            {
                                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TVA ERRONEE : " + VatNat + " " + Convert.ToString(VatPrc) + (char)(13) + (char)(10);
                                                FichierOk = false;
                                                LigneOk = false;
                                            }
                                        }
                                        else
                                        {
                                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TVA ERRONEE : " + VatNat + " " + Convert.ToString(VatPrc) + (char)(13) + (char)(10);
                                            FichierOk = false;
                                            LigneOk = false;
                                        }
                                    }
                                }
                                else
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TVA ERRONEE : " + VatNat + " " + Convert.ToString(VatPrc) + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                            else
                            {
                                if (Outils.Left(journalInfos.Dbtype, 1) != tvainfos.Vsalpur)
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TVA ERRONEE : " + VatNat + " " + Convert.ToString(VatPrc) + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                        }
                    }

                    // MONTANT BASE
                    if (!Outils.IsNumeric(ExcellArray[BIE_BASE], "DOUBLE"))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> MONTANT BASE INVALIDE : " + ExcellArray[BIE_BASE] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // MONTANT TVA
                    if (!Outils.IsNumeric(ExcellArray[BIE_TVA], "DOUBLE"))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> MONTANT TVA INVALIDE : " + ExcellArray[BIE_TVA] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // DEVISE
                    if (ExcellArray[BIE_DEVISE] != "" && ExcellArray[BIE_DEVISE] != DeviseBaseDossier)
                    {
                        Tuple<bool, bool> resultCurr = DaoBob.VerifStdCode("CURRENCY", ExcellArray[BIE_DEVISE], ReadBobIni.CommonDirectory);
                        if (!resultCurr.Item1)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DEVISE INVALIDE : " + ExcellArray[BIE_DEVISE] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            if (!resultCurr.Item2)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DEVISE INVISIBLE : " + ExcellArray[BIE_DEVISE] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }

                        // RECHERCHE D'UN COURS
                        if (!Outils.IsNumeric(ExcellArray[BIE_COURS], "DOUBLE"))
                        {
                            CoursDevise = bob50dao.GetCurrRate(folder, ExcellArray[BIE_DEVISE], DateDoc);

                            if (CoursDevise == (double)(0))
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COURS INVALIDE : " + ExcellArray[BIE_COURS] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                            else
                            {
                                if (CoursDevise <= 0)
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COURS INVALIDE : " + ExcellArray[BIE_COURS] + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                        }
                    }

                    if (!LigneOk)
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "------------------------------------------------------" + (char)(13) + (char)(10);

                    NumLigne = NumLigne + 1;

                    if (currForm.ProgressBar1.Value + 1 < 100)
                        currForm.ProgressBar1.Value = currForm.ProgressBar1.Value + 1;
                    else
                        currForm.ProgressBar1.Value = 0;
                }

                #endregion

                if (FichierOk)
                {
                    NumLigne = 1;
                    for (int r = 0; r < tableFichierExcel.Rows.Count; r++)
                    {

                        // LECTURE LIGNE EXCELL
                        ExcellArray.Clear();
                        for (int i = 0; i < NbColExcel; i++)
                        {
                            if (!DBNull.Value.Equals(tableFichierExcel.Rows[r][i]))
                                ExcellArray.Add(tableFichierExcel.Rows[r][i].ToString().Trim());
                            else
                                ExcellArray.Add("");
                        }

                        // POSITIONNEMENT SUR LE JOURNAL
                        Bob50Dbk journalInfos = DaoBob.GetJournalInfos(folder, ExcellArray[BIE_JOURNAL], folder.Path);

                        // POSITIONNEMENT SUR LA PERIODE POUR AVOIR L'EXERCICE FISCAL
                        Bob50Period periode = DaoBob.GetPeriode(ExcellArray[BIE_ANNEE], ExcellArray[BIE_MOIS], folder.Path);

                        // POSITIONNEMENT SUR LE CODE TVA BOB
                        if (TvaActive && journalInfos.Iswithvat)
                        {
                            VatType = "N";
                            if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                            {
                                Bob50Tiers customer = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "C", folder.Path);
                                if (customer.Cvatcat != "N")
                                {
                                    if (customer.Cvatref != LegisDossier)
                                    {
                                        cTemp = customer.Cvatref;
                                        switch (cTemp)
                                        {
                                            case "EX":
                                                VatType = "I";
                                                break;
                                            case "":
                                                VatType = "N";
                                                break;
                                            default:
                                                VatType = "E";
                                                break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Bob50Tiers supplier = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "S", folder.Path);
                                if (supplier.Cvatcat != "N")
                                {
                                    if (supplier.Cvatref != LegisDossier)
                                    {
                                        cTemp = supplier.Cvatref;
                                        switch (cTemp)
                                        {
                                            case "EX":
                                                VatType = "I";
                                                break;
                                            case "":
                                                VatType = "N";
                                                break;
                                            default:
                                                VatType = "E";
                                                break;
                                        }
                                    }
                                }
                            }
                            VatNat = ExcellArray[BIE_NATTVA].ToUpper();
                            VatPrc = Convert.ToDouble(ExcellArray[BIE_PRCTVA].Replace("%", ""));
                            tvainfos = DaoBob.GetVATInfos("S", VatType, VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                            if (tvainfos == null && VatType == "N" && (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC"))
                            {
                                tvainfos = DaoBob.GetVATInfos("S", "E", VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                                if (tvainfos == null && VatType == "N" && (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC"))
                                {
                                    tvainfos = DaoBob.GetVATInfos("S", "I", VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                                }
                            }
                        }

                        // CALCUL DU COURS DE LA DEVISE
                        if (!Outils.IsNumeric(ExcellArray[BIE_COURS], "DOUBLE"))
                        {
                            CoursDevise = (double)(0);

                            if (ExcellArray[BIE_DEVISE].Trim() != "")
                            {
                                CoursDevise = bob50dao.GetCurrRate(folder, ExcellArray[BIE_DEVISE], Convert.ToDateTime(ExcellArray[BIE_DATEDOC]));
                            }
                        }
                        else
                            CoursDevise = Convert.ToDouble(ExcellArray[BIE_COURS]);

                        // GENERATION DE L'ECRITURE
                        // ------------------------
                        if (RuptureDoc != ExcellArray[BIE_JOURNAL] + ExcellArray[BIE_NUMDOC])
                        {
                            RuptureDoc = ExcellArray[BIE_JOURNAL] + ExcellArray[BIE_NUMDOC];
                            IndImp = (int)(0);
                            TotalBase = (double)(0);
                            TotalTva = (double)(0);
                            TotalTTC = (double)(0);
                            Tva = (double)(0);

                            // ENTETE
                            cTemp = journalInfos.Dbtype;
                            switch (cTemp)
                            {
                                case "SAL":
                                    HeaderVEN.Tdocno = Convert.ToInt32(ExcellArray[BIE_NUMDOC]);
                                    HeaderVEN.Tdbk = journalInfos.Dbid;
                                    HeaderVEN.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                    HeaderVEN.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                    HeaderVEN.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    HeaderVEN.TtypeCie = "C";
                                    HeaderVEN.TCompan = ExcellArray[BIE_TIERS];
                                    NumDocVEN = Convert.ToInt32(ExcellArray[BIE_NUMDOC]);
                                    HeaderVEN.Tfyear = periode.Fyear;
                                    if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                        HeaderVEN.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEECH]);
                                    else
                                        HeaderVEN.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                        HeaderVEN.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    else
                                    {
                                        TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                        HeaderVEN.TAmount = TotalTTC;
                                        HeaderVEN.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                        HeaderVEN.TCurrency = ExcellArray[BIE_DEVISE];
                                        HeaderVEN.TCurRate = CoursDevise;
                                    }
                                    HeaderVEN.TRemExt = ExcellArray[BIE_COMM];
                                    HeaderVEN.TRemInt = ExcellArray[BIE_REM];
                                    HeaderVEN.Tintmode = "S";
                                    if (ExcellArray[BIE_NAMEDOC] != "")
                                        HeaderVEN.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);
                                    DaoBob.AddEntete(HeaderVEN, VarGlobal.BOBDossierRun, "VEN");
                                    break;
                                case "SAC":
                                    HeaderNCV.Tdocno = Convert.ToInt32(ExcellArray[BIE_NUMDOC]);
                                    HeaderNCV.Tdbk = journalInfos.Dbid;
                                    HeaderNCV.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                    HeaderNCV.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                    HeaderNCV.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    HeaderNCV.TtypeCie = "C";
                                    HeaderNCV.TCompan = ExcellArray[BIE_TIERS];
                                    HeaderNCV.Tfyear = periode.Fyear;
                                    NumDocNCV = Convert.ToInt32(ExcellArray[BIE_NUMDOC]);
                                    if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                        HeaderNCV.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    else
                                        HeaderNCV.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                        HeaderNCV.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    else
                                    {
                                        TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                        HeaderNCV.TAmount = TotalTTC;
                                        HeaderNCV.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                        HeaderNCV.TCurrency = ExcellArray[BIE_DEVISE];
                                        HeaderNCV.TCurRate = CoursDevise;
                                    }
                                    HeaderNCV.TRemExt = ExcellArray[BIE_COMM];
                                    HeaderNCV.TRemInt = ExcellArray[BIE_REM];
                                    HeaderNCV.Tintmode = "S";
                                    if (ExcellArray[BIE_NAMEDOC] != "")
                                        HeaderNCV.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);
                                    DaoBob.AddEntete(HeaderNCV, VarGlobal.BOBDossierRun, "NCV");
                                    break;
                                case "PUR":
                                    NumDocACH = NumDocACH + 1;
                                    HeaderACH.Tdocno = Convert.ToInt32(NumDocACH);
                                    HeaderACH.Tdbk = journalInfos.Dbid;
                                    HeaderACH.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                    HeaderACH.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                    HeaderACH.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    HeaderACH.TtypeCie = "S";
                                    HeaderACH.TCompan = ExcellArray[BIE_TIERS];
                                    HeaderACH.Tfyear = periode.Fyear;
                                    if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                        HeaderACH.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEECH]);
                                    else
                                        HeaderACH.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                        HeaderACH.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    else
                                    {
                                        TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                        HeaderACH.TAmount = TotalTTC;
                                        HeaderACH.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                        HeaderACH.TCurrency = ExcellArray[BIE_DEVISE];
                                        HeaderACH.TCurRate = CoursDevise;
                                    }
                                    HeaderACH.TRemExt = ExcellArray[BIE_COMM];
                                    HeaderACH.TRemInt = ExcellArray[BIE_REM];
                                    HeaderACH.Tintmode = "S";
                                    if (ExcellArray[BIE_NAMEDOC] != "")
                                        HeaderACH.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);

                                    DaoBob.AddEntete(HeaderACH, VarGlobal.BOBDossierRun, "ACH");
                                    break;
                                case "PUC":
                                    NumDocNCA = NumDocNCA + 1;
                                    HeaderNCA.Tdocno = Convert.ToInt32(NumDocNCA);
                                    HeaderNCA.Tdbk = journalInfos.Dbid;
                                    HeaderNCA.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                    HeaderNCA.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                    HeaderNCA.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    HeaderNCA.TtypeCie = "S";
                                    HeaderNCA.TCompan = ExcellArray[BIE_TIERS];
                                    HeaderNCA.Tfyear = periode.Fyear;
                                    if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                        HeaderNCA.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    else
                                        HeaderNCA.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                    if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                        HeaderNCA.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    else
                                    {
                                        TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                        HeaderNCA.TAmount = TotalTTC;
                                        HeaderNCA.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                        HeaderNCA.TCurrency = ExcellArray[BIE_DEVISE];
                                        HeaderNCA.TCurRate = CoursDevise;
                                    }
                                    HeaderNCA.TRemExt = ExcellArray[BIE_COMM];
                                    HeaderNCA.TRemInt = ExcellArray[BIE_REM];
                                    HeaderNCA.Tintmode = "S";
                                    if (ExcellArray[BIE_NAMEDOC] != "")
                                        HeaderNCA.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);

                                    DaoBob.AddEntete(HeaderNCA, VarGlobal.BOBDossierRun, "NCA");
                                    break;
                            }
                        }

                        // IMPUTATION
                        IndImp = IndImp + (int)(1);
                        cTemp = journalInfos.Dbtype;
                        switch (cTemp)
                        {
                            case "SAL":
                                LineVEN.Tdocno = NumDocVEN;
                                LineVEN.Tdbk = journalInfos.Dbid;
                                LineVEN.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                LineVEN.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                LineVEN.Tdocline = IndImp;
                                LineVEN.Ttypeline = "S";
                                LineVEN.Tacttype = "";
                                LineVEN.Taccount = ExcellArray[BIE_COMPTE];
                                LineVEN.Tfyear = periode.Fyear;
                                LineVEN.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                {
                                    BaseTva = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    Tva = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TotalBase = TotalBase + BaseTva;
                                    TotalTva = TotalTva + Tva;
                                    LineVEN.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineVEN.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineVEN.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineVEN.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineVEN.TVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineVEN.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                else
                                {
                                    BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                    Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                    TotalBase = TotalBase + BaseTva;
                                    TotalTva = TotalTva + Tva;
                                    LineVEN.Tamount = BaseTva;
                                    LineVEN.TBasVat = BaseTva;
                                    LineVEN.TBasLstAmn = BaseTva;
                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineVEN.TVatTotAmn = Tva;
                                        TvaDeduc = Math.Round(Tva / 100 * tvainfos.Vdeduc, 2);
                                        LineVEN.TVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineVEN.TVcDblAmn = Tva;
                                    }

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {

                                        LineVEN.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                        LineVEN.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                        LineVEN.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineVEN.TCurVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineVEN.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                if (TvaActive && journalInfos.Iswithvat)
                                    LineVEN.TVStored = tvainfos.Vstored;
                                LineVEN.Tdc = "C";
                                LineVEN.TRem = ExcellArray[BIE_REM];

                                // ANALYTIQUE
                                for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            LineVEN.Cost_1 = ExcellArray[NbColStd + 1];
                                            break;
                                        case 1:
                                            LineVEN.Cost_2 = ExcellArray[NbColStd + 2];
                                            break;
                                        case 2:
                                            LineVEN.Cost_3 = ExcellArray[NbColStd + 3];
                                            break;
                                        case 3:
                                            LineVEN.Cost_4 = ExcellArray[NbColStd + 4];
                                            break;
                                        case 4:
                                            LineVEN.Cost_5 = ExcellArray[NbColStd + 5];
                                            break;
                                        case 5:
                                            LineVEN.Cost_6 = ExcellArray[NbColStd + 6];
                                            break;
                                        case 6:
                                            LineVEN.Cost_7 = ExcellArray[NbColStd + 7];
                                            break;
                                        case 7:
                                            LineVEN.Cost_8 = ExcellArray[NbColStd + 8];
                                            break;
                                    }
                                }

                                DaoBob.AddLigne(LineVEN, ListeSecAna, VarGlobal.BOBDossierRun, "VEN");
                                break;
                            case "SAC":
                                LineNCV.Tdocno = NumDocNCV;
                                LineNCV.Tdbk = journalInfos.Dbid;
                                LineNCV.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                LineNCV.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                LineNCV.Tdocline = IndImp;
                                LineNCV.Ttypeline = "S";
                                LineNCV.Tacttype = "";
                                LineNCV.Taccount = ExcellArray[BIE_COMPTE];
                                LineNCV.Tfyear = periode.Fyear;
                                LineNCV.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                {
                                    LineNCV.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineNCV.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineNCV.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineNCV.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineNCV.TVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineNCV.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                else
                                {
                                    BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                    Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                    TotalBase = TotalBase + BaseTva;
                                    TotalTva = TotalTva + Tva;
                                    LineNCV.Tamount = BaseTva;
                                    LineNCV.TBasVat = BaseTva;
                                    LineNCV.TBasLstAmn = BaseTva;

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineNCV.TVatTotAmn = Tva;
                                        LineNCV.TVatAmn = Tva;
                                        if (tvainfos.Vkind == "D")
                                            LineNCV.TVatDblAmn = Tva;
                                    }

                                    LineNCV.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineNCV.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineNCV.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineNCV.TCurVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineNCV.TVcDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                if (TvaActive && journalInfos.Iswithvat)
                                    LineNCV.TVStored = tvainfos.Vstored;
                                LineNCV.Tdc = "D";
                                LineNCV.TRem = ExcellArray[BIE_REM];

                                // ANALYTIQUE
                                for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            LineNCV.Cost_1 = ExcellArray[NbColStd + 1];
                                            break;
                                        case 1:
                                            LineNCV.Cost_2 = ExcellArray[NbColStd + 2];
                                            break;
                                        case 2:
                                            LineNCV.Cost_3 = ExcellArray[NbColStd + 3];
                                            break;
                                        case 3:
                                            LineNCV.Cost_4 = ExcellArray[NbColStd + 4];
                                            break;
                                        case 4:
                                            LineNCV.Cost_5 = ExcellArray[NbColStd + 5];
                                            break;
                                        case 5:
                                            LineNCV.Cost_6 = ExcellArray[NbColStd + 6];
                                            break;
                                        case 6:
                                            LineNCV.Cost_7 = ExcellArray[NbColStd + 7];
                                            break;
                                        case 7:
                                            LineNCV.Cost_8 = ExcellArray[NbColStd + 8];
                                            break;
                                    }
                                }

                                DaoBob.AddLigne(LineNCV, ListeSecAna, VarGlobal.BOBDossierRun, "NCV");
                                break;
                            case "PUR":
                                LineACH.Tdocno = NumDocACH;
                                LineACH.Tdbk = journalInfos.Dbid;
                                LineACH.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                LineACH.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                LineACH.Tdocline = IndImp;
                                LineACH.Ttypeline = "S";
                                LineACH.Tacttype = "";
                                LineACH.Taccount = ExcellArray[BIE_COMPTE];
                                LineACH.Tfyear = periode.Fyear;
                                LineACH.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                {
                                    LineACH.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineACH.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineACH.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineACH.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineACH.TVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineACH.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                else
                                {
                                    BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                    Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                    TotalBase = TotalBase + BaseTva;
                                    TotalTva = TotalTva + Tva;
                                    LineACH.Tamount = BaseTva;
                                    LineACH.TBasVat = BaseTva;
                                    LineACH.TBasLstAmn = BaseTva;

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineACH.TVatTotAmn = Tva;
                                        LineACH.TVatAmn = Tva;
                                        if (tvainfos.Vkind == "D")
                                            LineACH.TVatDblAmn = Tva;
                                    }

                                    LineACH.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineACH.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineACH.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineACH.TCurVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineACH.TVcDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                if (TvaActive && journalInfos.Iswithvat)
                                    LineACH.TVStored = tvainfos.Vstored;
                                LineACH.Tdc = "D";
                                LineACH.TRem = ExcellArray[BIE_REM];

                                // ANALYTIQUE
                                for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            LineACH.Cost_1 = ExcellArray[NbColStd + 1];
                                            break;
                                        case 1:
                                            LineACH.Cost_2 = ExcellArray[NbColStd + 2];
                                            break;
                                        case 2:
                                            LineACH.Cost_3 = ExcellArray[NbColStd + 3];
                                            break;
                                        case 3:
                                            LineACH.Cost_4 = ExcellArray[NbColStd + 4];
                                            break;
                                        case 4:
                                            LineACH.Cost_5 = ExcellArray[NbColStd + 5];
                                            break;
                                        case 5:
                                            LineACH.Cost_6 = ExcellArray[NbColStd + 6];
                                            break;
                                        case 6:
                                            LineACH.Cost_7 = ExcellArray[NbColStd + 7];
                                            break;
                                        case 7:
                                            LineACH.Cost_8 = ExcellArray[NbColStd + 8];
                                            break;
                                    }
                                }

                                DaoBob.AddLigne(LineACH, ListeSecAna, VarGlobal.BOBDossierRun, "ACH");
                                break;
                            case "PUC":
                                LineNCA.Tdocno = NumDocNCA;
                                LineNCA.Tdbk = journalInfos.Dbid;
                                LineNCA.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                LineNCA.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                LineNCA.Tdocline = IndImp;
                                LineNCA.Ttypeline = "S";
                                LineNCA.Tacttype = "";
                                LineNCA.Taccount = ExcellArray[BIE_COMPTE];
                                LineNCA.Tfyear = periode.Fyear;
                                LineNCA.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                {
                                    LineNCA.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineNCA.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineNCA.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                        LineNCA.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineNCA.TVatAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineNCA.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                else
                                {
                                    BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                    Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                    TotalBase = TotalBase + BaseTva;
                                    TotalTva = TotalTva + Tva;
                                    LineNCA.Tamount = BaseTva;
                                    LineNCA.TBasVat = BaseTva;
                                    LineNCA.TBasLstAmn = BaseTva;

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineNCA.TVatTotAmn = Tva;
                                        LineNCA.TVatAmn = Tva;
                                        if (tvainfos.Vkind == "D")
                                            LineNCA.TVatDblAmn = Tva;
                                    }

                                    LineNCA.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineNCA.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);

                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        LineNCA.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                        TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                        LineNCA.TCurAmn = TvaDeduc;
                                        if (tvainfos.Vkind == "D")
                                            LineNCA.TVcDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    }
                                }
                                if (TvaActive && journalInfos.Iswithvat)
                                    LineNCA.TVStored = tvainfos.Vstored;
                                LineNCA.Tdc = "C";
                                LineNCA.TRem = ExcellArray[BIE_REM];

                                // ANALYTIQUE
                                for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            LineNCA.Cost_1 = ExcellArray[NbColStd + 1];
                                            break;
                                        case 1:
                                            LineNCA.Cost_2 = ExcellArray[NbColStd + 2];
                                            break;
                                        case 2:
                                            LineNCA.Cost_3 = ExcellArray[NbColStd + 3];
                                            break;
                                        case 3:
                                            LineNCA.Cost_4 = ExcellArray[NbColStd + 4];
                                            break;
                                        case 4:
                                            LineNCA.Cost_5 = ExcellArray[NbColStd + 5];
                                            break;
                                        case 5:
                                            LineNCA.Cost_6 = ExcellArray[NbColStd + 6];
                                            break;
                                        case 6:
                                            LineNCA.Cost_7 = ExcellArray[NbColStd + 7];
                                            break;
                                        case 7:
                                            LineNCA.Cost_8 = ExcellArray[NbColStd + 8];
                                            break;
                                    }
                                }

                                DaoBob.AddLigne(LineNCA, ListeSecAna, VarGlobal.BOBDossierRun, "NCA");
                                break;
                        }
                        
                        NumLigne = NumLigne + 1;

                        bool rupture = false;
                        if (r + 1 < tableFichierExcel.Rows.Count)
                        {
                            cTemp = tableFichierExcel.Rows[r + 1][0].ToString().Trim() + tableFichierExcel.Rows[r + 1][5].ToString().Trim();
                            if (RuptureDoc != cTemp)
                                rupture = true;
                        }
                        else
                        {
                            rupture = true;
                        }

                        // CONTROLES DE FIN DE RUPTURE
                        // if (r+1 >= tableFichierExcel.Rows.Count || RuptureDoc != ValCol1Excel + cTemp)
                        if (rupture)
                        {
                            if (ExcellArray[BIE_DEVISE] != "" && ExcellArray[BIE_DEVISE] != DeviseBaseDossier)
                            {
                                Ecart = Math.Round(TotalTTC - Math.Round(TotalBase + TotalTva, 2), 2);
                                if (Ecart != (double)(0))
                                {
                                    if (Tva == (double)(0))
                                    {
                                        cTemp = journalInfos.Dbtype;
                                        switch (cTemp)
                                        {
                                            case "SAL":
                                                BaseTva = Math.Round(Convert.ToDouble(LineVEN.Tamount) + Ecart, 2);
                                                LineVEN.Tamount = BaseTva;
                                                LineVEN.TBasVat = BaseTva;
                                                LineVEN.TBasLstAmn = BaseTva;
                                                DaoBob.UpdateLigneBASE(LineVEN, VarGlobal.BOBDossierRun, "VEN");
                                                break;
                                            case "SAC":
                                                BaseTva = Math.Round(Convert.ToDouble(LineNCV.Tamount) + Ecart, 2);
                                                LineNCV.Tamount = BaseTva;
                                                LineNCV.TBasVat = BaseTva;
                                                LineNCV.TBasLstAmn = BaseTva;
                                                DaoBob.UpdateLigneBASE(LineNCV, VarGlobal.BOBDossierRun, "NCV");
                                                break;
                                            case "PUR":
                                                BaseTva = Math.Round(Convert.ToDouble(LineACH.Tamount) + Ecart, 2);
                                                LineACH.Tamount = BaseTva;
                                                LineACH.TBasVat = BaseTva;
                                                LineACH.TBasLstAmn = BaseTva;
                                                DaoBob.UpdateLigneBASE(LineACH, VarGlobal.BOBDossierRun, "ACH");
                                                Log.WriteLog("Ecart", "BASETVA;" + BaseTva.ToString() + ";" + Ecart.ToString());
                                                break;
                                            case "PUC":
                                                BaseTva = Math.Round(Convert.ToDouble(LineNCA.Tamount) + Ecart, 2);
                                                LineNCA.Tamount = BaseTva;
                                                LineNCA.TBasVat = BaseTva;
                                                LineNCA.TBasLstAmn = BaseTva;
                                                DaoBob.UpdateLigneBASE(LineNCA, VarGlobal.BOBDossierRun, "NCA");
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        if (TvaActive && journalInfos.Iswithvat)
                                        {
                                            cTemp = journalInfos.Dbtype;
                                            switch (cTemp)
                                            {
                                                case "SAL":
                                                    Tva = Math.Round(Convert.ToDouble(LineVEN.TVatAmn) + Ecart, 2);
                                                    LineVEN.TVatAmn = Tva;
                                                    LineVEN.TVatTotAmn = Tva;
                                                    DaoBob.UpdateLigneTVA(LineVEN, VarGlobal.BOBDossierRun, "VEN");
                                                    break;
                                                case "SAC":
                                                    Tva = Math.Round(Convert.ToDouble(LineNCV.TVatAmn) + Ecart, 2);
                                                    LineNCV.TVatAmn = Tva;
                                                    LineNCV.TVatTotAmn = Tva;
                                                    DaoBob.UpdateLigneTVA(LineNCV, VarGlobal.BOBDossierRun, "NCV");
                                                    break;
                                                case "PUR":
                                                    Tva = Math.Round(Convert.ToDouble(LineACH.TVatAmn) + Ecart, 2);
                                                    LineACH.TVatAmn = Tva;
                                                    LineACH.TVatTotAmn = Tva;
                                                    DaoBob.UpdateLigneTVA(LineACH, VarGlobal.BOBDossierRun, "ACH");
                                                    Log.WriteLog("Ecart", "TVA;" + Tva.ToString() + ";" + Ecart.ToString());
                                                    break;
                                                case "PUC":
                                                    Tva = Math.Round(Convert.ToDouble(LineNCA.TVatAmn) + Ecart, 2);
                                                    LineNCA.TVatAmn = Tva;
                                                    LineNCA.TVatTotAmn = Tva;
                                                    DaoBob.UpdateLigneTVA(LineNCA, VarGlobal.BOBDossierRun, "NCA");
                                                    break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (currForm.ProgressBar1.Value + 1 < 100)
                        currForm.ProgressBar1.Value = currForm.ProgressBar1.Value + 1;
                    else
                        currForm.ProgressBar1.Value = 0;
                }

                // REINITIALISATION PROGRESS BAR
                currForm.ProgressBar1.Value = (int)(1);

                // FICHIER EN ERREUR
                if (!FichierOk)
                {
                    MessageBox.Show("Erreurs dans le fichier d'import, veuillez corriger", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                LinkOk = true;

                // LIAISON DES VENTES
                if (NumDocVEN > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHVEN.ADT /LNF=LINK_KLVEN.ADT /SEPARATE /NOPOST /MOD=ENTRYSAL /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG")))
                    {
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                // LIAISON DES NOTES DE CREDIT
                if (NumDocNCV > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHNCV.ADT /LNF=LINK_KLNCV.ADT /SEPARATE /NOPOST /MOD=ENTRYSAL /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG")))
                    {
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                // LIAISON DES ACHATS
                if (NumDocACH > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHACH.ADT /LNF=LINK_KLACH.ADT /SEPARATE /NOPOST /MOD=ENTRYPUR /AUTONUM /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG")))
                    {
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();                       
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                // LIAISON DES NOTES DE CREDIT SUR ACHAT
                if (NumDocNCA > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHNCA.ADT /LNF=LINK_KLNCA.ADT /SEPARATE /NOPOST /MOD=ENTRYPUR /AUTONUM /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG")))
                    {
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, VarGlobal.BOBDossierRun, "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                currForm.InfosImport.Text = currForm.InfosImport.Text;

                // LIAISON REELLE ET ARCHIVAGE DU FICHIER XLS IMPORTE
                if (LinkOk)
                {
                    VarGlobal.listPdf.Clear();

                    // LIAISON DES VENTES
                    if (NumDocVEN > (long)(0))
                    {
                        if (currForm.cb_NumAutoVentes.Checked)
                            RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHVEN.ADT /LNF=LINK_KLVEN.ADT /INTEMP /MOD=ENTRYSAL /AUTONUM /AUTOCLOSE");
                        else
                            RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHVEN.ADT /LNF=LINK_KLVEN.ADT /INTEMP /MOD=ENTRYSAL /AUTOCLOSE");

                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "VEN");
                    }

                    // LIAISON DES NOTES DE CREDIT
                    if (NumDocNCV > (long)(0))
                    {
                        if (currForm.cb_NumAutoVentes.Checked)
                            RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHNCV.ADT /LNF=LINK_KLNCV.ADT /INTEMP /MOD=ENTRYSAL /AUTONUM /AUTOCLOSE");
                        else
                            RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHNCV.ADT /LNF=LINK_KLNCV.ADT /INTEMP /MOD=ENTRYSAL /AUTOCLOSE");

                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "NCV");
                    }

                    // LIAISON DES ACHATS
                    if (NumDocACH > (long)(0))
                    {
                        RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHACH.ADT /LNF=LINK_KLACH.ADT /INTEMP /MOD=ENTRYPUR /AUTONUM /AUTOCLOSE");

                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "ACH");
                    }

                    // LIAISON DES NOTES DE CREDIT SUR ACHAT
                    if (NumDocNCA > (long)(0))
                    {
                        RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + VarGlobal.BOBDossierRun + " /HDF=LINK_KHNCA.ADT /LNF=LINK_KLNCA.ADT /INTEMP /MOD=ENTRYPUR /AUTONUM /AUTOCLOSE");

                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "NCA");
                    }

                    // ARCHIVAGE DES PDF
                    if (!Directory.Exists(Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], "Archives")))
                        Directory.CreateDirectory(Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], "Archives"));
                    foreach (string pdf in VarGlobal.listPdf)
                    {
                        if (File.Exists(pdf))
                        {
                            File.Move(pdf, Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], "Archives", Path.GetFileName(pdf)));
                        }
                    }


                    currForm.FichierImport.Text = "";
                    MessageBox.Show("Import terminé", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Problèmes lors de l'import, contacter votre revendeur", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception ex)
            {
                Log.WriteLog("ImportExcel", "PARTIE DOSSIER SIMPLE : " + ex.Message);
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        /// <summary>
        /// IMPORT EXCEL EN MODE MULTI DOSSIERS
        /// </summary>
        public void RunImportMultiDossiers()
        {
            const int BIE_DOSSIER = 0;
            const int BIE_JOURNAL = 1;
            const int BIE_ANNEE = 2;
            const int BIE_MOIS = 3;
            const int BIE_DATEDOC = 4;
            const int BIE_DATEECH = 5;
            const int BIE_NUMDOC = 6;
            const int BIE_TIERS = 7;
            const int BIE_COMPTE = 8;
            const int BIE_REM = 9;
            const int BIE_COMM = 10;
            const int BIE_TOTALTTC = 11;
            const int BIE_NATTVA = 12;
            const int BIE_PRCTVA = 13;
            const int BIE_BASE = 14;
            const int BIE_TVA = 15;
            const int BIE_DEVISE = 16;
            const int BIE_COURS = 17;
            const int BIE_TIERSNOM = 18;
            const int BIE_TIERSADR = 19;
            const int BIE_TIERSCP = 20;
            const int BIE_TIERSVILLE = 21;
            const int BIE_TIERSASS = 22;
            const int BIE_TIERSTVA = 23;
            const int BIE_TIERSEMAIL = 24;
            const int BIE_NAMEDOC = 25;

            int NbColExcel;
            List<string> ListeDossiers = new List<string>();
            List<string> ListeDossiersPath = new List<string>();
            List<string> ListeSecAna = new List<string>();
            List<Boolean> ListeSecAnaMandatory = new List<Boolean>();

            Bob50KhDbk HeaderACH = new Bob50KhDbk();
            Bob50KhDbk HeaderNCA = new Bob50KhDbk();
            Bob50KhDbk HeaderVEN = new Bob50KhDbk();
            Bob50KhDbk HeaderNCV = new Bob50KhDbk();

            Bob50KlDbk LineACH = new Bob50KlDbk();
            Bob50KlDbk LineNCA = new Bob50KlDbk();
            Bob50KlDbk LineVEN = new Bob50KlDbk();
            Bob50KlDbk LineNCV = new Bob50KlDbk();

            int NumLigne;

            List<DmDoc> listDmDoc = new List<DmDoc>();
            List<DmInvDoc> listDmInvDoc = new List<DmInvDoc>();

            string RuptureDossier = "";
            string RuptureDoc = "";
            int IndImp = 0;
            double BaseTva = 0;
            double Tva = 0;
            double TvaDeduc = 0;
            double TotalTTC = 0;
            double TotalBase = 0;
            double TotalTva = 0;
            double Ecart = 0;
            bool FichierOk;
            bool LigneOk;
            string VatType = "";
            string VatNat = "";
            Double VatPrc = 0;
            double CoursDevise = 0;
            DateTime DateDoc;
            DateTime DateEch;
            string DeviseBaseDossier = "";
            String LegisDossier = "";
            TVAInfos tvainfos = null;
            int NbColStd = 25;
            Bob50Period periode = null;

            bool LinkOk;
            var RetVal = 0;
            string cTemp = "";

            Boolean AnaObligatoireGlobale = false;
            Boolean TvaActive = true;

            // ARRAY POUR LECTURE LIGNE EXCELL
            List<string> ExcellArray = new List<string>();

            // TEST SI FICHIER D'IMPORT BIEN SELECTIONNE
            if (String.IsNullOrEmpty(currForm.FichierImport.Text))
            {
                MessageBox.Show("Veuillez d'abord sélectionner un fichier Excel à importer", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!File.Exists(currForm.FichierImport.Text))
            {
                MessageBox.Show("Fichier d'import inexistant", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // CONFIRMATION DE LA LIAISON
            if (MessageBox.Show("Démarrer l'import", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
                return;

            RuptureDoc = "";

            // CREATION DU REPERTOIRE LINK DU DOSSIER SI NECESSAIRE ET COPIE DES FICHIERS DE LIAISON
            Bob50Dao bob50dao = new Bob50Dao();

            // LECTURE DU FICHIER EXCEL
            DataTable tableFichierExcel = Dev4LogLayerFiles.Excel.ExcelToDataTable(currForm.FichierImport.Text, null, true, true, 0, true, null, 0, null, false, false);

            #region INITIALISATION NOMBRE DE COLONNES ET DES CODES DES SECTIONS ANALYTIQUES

            NbColExcel = 25;
            for (int i = 28; i <= 34; i++)
            {
                if (i < tableFichierExcel.Columns.Count)
                {
                    // cTemp = (range.Cells[1, i] as Excel.Range).Text;
                    cTemp = tableFichierExcel.Rows[0][i].ToString();
                    if (cTemp != null)
                    {
                        if (cTemp.Trim() != "")
                        {
                            NbColExcel = NbColExcel + 1;
                            ListeSecAna.Add(cTemp);
                        }
                        else
                            break;
                    }
                    else
                        break;
                }
            }

            #endregion


            currForm.ProgressBar1.Value = 01;
            currForm.ProgressBar1.Visible = true;
            FichierOk = true;

            // INITIALISATIONS
            int NumDocACH = 0;
            int NumDocNCA = 0;
            int NumDocVEN = 0;
            int NumDocNCV = 0;

            // PARCOURS DU FICHIER EXCEL
            currForm.InfosImport.Text = "";
            RuptureDoc = "";

            string dossierActuel = "";
            Bob50Folder folder = null;
            Bob50Config conf = null;


            #region LECTURE ET CONTROLE DE COHERENCE DU FICHIER EXCEL

            NumLigne = 2;
            foreach (DataRow dr in tableFichierExcel.Rows)
            {
                if (dr[0].ToString() == "")
                    continue;

                // LECTURE LIGNE EXCELL
                ExcellArray.Clear();
                for (int i = 0; i <= NbColExcel; i++)
                {
                    if (i == 13 || i == 14 || i == 15 || i == 17)
                    {
                        if (!String.IsNullOrEmpty(dr[i].ToString()))
                        {
                            cTemp = Convert.ToString(Convert.ToDouble(dr[i]));
                        }
                    }
                    else
                        cTemp = dr[i].ToString();
                    if (cTemp != null)
                        ExcellArray.Add(cTemp.Trim());
                    else
                        ExcellArray.Add("");
                }

                // CONNEXION DOSSIER BOB ET INITIALISATIONS
                // ----------------------------------------
                LigneOk = true;
                if (RuptureDossier != dr[BIE_DOSSIER].ToString())
                {

                    // CONNEXTION DOSSIER BOB ET OUVERTURE DES TABLES
                    cTemp = dossierActuel;
                    dossierActuel = dr[BIE_DOSSIER].ToString();
                    folder = bob50dao.GetBobFolder(dossierActuel);
                    conf = DaoBob.LOCAL_GetConfigDossier(folder);
                    DeviseBaseDossier = conf.Basecurrid;
                    LegisDossier = conf.Legis;
                    AnaObligatoireGlobale = conf.Costmand;
                    TvaActive = conf.Prmvatenbld;
                    if (folder.Id == dr[BIE_DOSSIER].ToString())
                    {
                        RuptureDossier = dr[BIE_DOSSIER].ToString();
                        ListeDossiers.Add(folder.Id);
                        ListeDossiersPath.Add(folder.Path);
                        if (!ListeDossiersPath[ListeDossiersPath.Count - 1].EndsWith(@"\"))
                            ListeDossiersPath[ListeDossiersPath.Count - 1] = ListeDossiersPath[ListeDossiersPath.Count - 1] + @"\";

                        // CONTROLE DES SECTIONS ANALYTIQUES ET INITALISATION PARTIE OBLIGATOIRE
                        ListeSecAnaMandatory.Clear();
                        for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                        {
                            string sectionObligatoire = DaoBob.GetSecAnaObligatoire(ListeSecAna[i], folder.Path);
                            switch (sectionObligatoire)
                            {
                                case "TRUE":
                                    ListeSecAnaMandatory.Add(true);
                                    break;

                                case "FALSE":
                                    ListeSecAnaMandatory.Add(false);
                                    break;

                                default:
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "SECTION ANALYTIQUE INEXISTANTE : " + ListeSecAna[i] + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    break;
                            }
                            if (FichierOk == false)
                                break;
                        }

                    }
                    else
                    {

                        // DOSSIER INEXISTANT -> REOUVERTURE DOSSIER DE DEPART
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DOSSIER INEXISTANT : " + dr[0].ToString() + (char)(13) + (char)(10);
                        dossierActuel = cTemp;
                        LigneOk = false;
                        FichierOk = false;
                    }
                }

                // CONTROLES ET AJOUT/MODIFICATIONS PREALABLES
                // -------------------------------------------

                Bob50Dbk journalInfos = new Bob50Dbk();

                if (LigneOk)
                {
                    // JOURNAL
                    if (ExcellArray[BIE_JOURNAL] == "")
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> JOURNAL VIDE" + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                        continue;
                    }
                    if (!DaoBob.GetJournalOk(ExcellArray[BIE_JOURNAL], folder.Path))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> JOURNAL INEXISTANT : " + ExcellArray[BIE_JOURNAL] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                        continue;
                    }
                    else
                    {                       
                        journalInfos = DaoBob.GetJournalInfos(folder, ExcellArray[BIE_JOURNAL], folder.Path);
                    }

                    // PERIODE
                    periode = DaoBob.GetPeriode(ExcellArray[BIE_ANNEE], ExcellArray[BIE_MOIS], folder.Path);
                    if (periode.Month == 0 || periode.Year == 0)
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> PERIODE INEXISTANTE : " + ExcellArray[BIE_MOIS] + " " + ExcellArray[BIE_ANNEE] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    if (ExcellArray[BIE_NAMEDOC] != "" && !File.Exists(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"] + ExcellArray[BIE_NAMEDOC]))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> Le document pdf entré n'existe pas." + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // DATE
                    if (!Outils.IsDate(ExcellArray[BIE_DATEDOC]))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE INVALIDE : " + ExcellArray[BIE_DATEDOC] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                        DateDoc = DateTime.Now;
                    }
                    else
                    {
                        DateDoc = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                        if (DaoBob.GetJournalOk(ExcellArray[BIE_JOURNAL], folder.Path))
                        {
                            if ((journalInfos.Dbtype == "SAL") || (journalInfos.Dbtype == "SAC"))
                            {
                                if (DateDoc.Month != periode.Month || DateDoc.Year != periode.Year)
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE N'APPARTIENT PAS A LA PERIODE : " + ExcellArray[BIE_DATEDOC] + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                            else
                            {
                                if (DateDoc.Month > periode.Month && DateDoc.Year == periode.Year)
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE SUPERIEURE A LA PERIODE : " + ExcellArray[BIE_DATEDOC] + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                        }
                    }

                    // DATE ECHEANCE
                    if (ExcellArray[BIE_DATEECH] == "")
                        DateEch = DateDoc;
                    else
                    {
                        if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                        {
                            DateEch = Convert.ToDateTime(ExcellArray[BIE_DATEECH]);
                            if (DateEch < DateDoc)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE D'ECHEANCE INFERIEURE A LA DATE DU DOCUMENT : " + ExcellArray[BIE_DATEECH] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }
                        else
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DATE D'ECHEANCE INVALIDE : " + ExcellArray[BIE_DATEECH] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                            DateEch = DateTime.Now;
                        }
                    }

                    // AJOUT MODIFICATION TIERS SI NECESSAIRE
                    if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSNOM]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSADR]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSCP]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSVILLE]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSTVA]) || !string.IsNullOrEmpty(ExcellArray[BIE_TIERSASS]))
                    {
                        Tuple<bool, bool> resultTva = DaoBob.VerifStdCode("VATCOUNT", ExcellArray[BIE_TIERSASS], ReadBobIni.CommonDirectory);
                        Bob50Tiers ajoutCompan = new Bob50Tiers();
                        if (ExcellArray[BIE_TIERSASS] == "EX" || ExcellArray[BIE_TIERSASS] == "NA" || resultTva.Item1)
                        {
                            ajoutCompan.Cid = ExcellArray[BIE_TIERS];
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSNOM]))
                                ajoutCompan.Cname1 = ExcellArray[BIE_TIERSNOM].Replace("'", "''");
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSADR]))
                                ajoutCompan.Cadress1 = ExcellArray[BIE_TIERSADR].Replace("'", "''");
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSCP]))
                                ajoutCompan.Czipcode = ExcellArray[BIE_TIERSCP];
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSVILLE]))
                                ajoutCompan.Clocality = ExcellArray[BIE_TIERSVILLE];
                            if (!string.IsNullOrEmpty(ExcellArray[BIE_TIERSEMAIL]))
                                ajoutCompan.Emailaddress = ExcellArray[BIE_TIERSEMAIL];
                            switch (ExcellArray[BIE_TIERSASS])
                            {
                                case "NA":
                                    ajoutCompan.Cvatcat = "N";
                                    ajoutCompan.Cvatref = "";
                                    ajoutCompan.Cvatno = "";
                                    break;
                                case "EX":
                                    ajoutCompan.Cvatcat = "";
                                    ajoutCompan.Cvatref = "EX";
                                    ajoutCompan.Cvatno = ExcellArray[BIE_TIERSTVA];
                                    break;
                                default:
                                    ajoutCompan.Cvatcat = "";
                                    ajoutCompan.Cvatref = ExcellArray[BIE_TIERSASS];
                                    ajoutCompan.Cvatno = ExcellArray[BIE_TIERSTVA];
                                    break;
                            }

                            if (DaoBob.GetJournalOk(ExcellArray[BIE_JOURNAL], folder.Path))
                                journalInfos = DaoBob.GetJournalInfos(folder, ExcellArray[BIE_JOURNAL], folder.Path);
                            else
                                journalInfos = new Bob50Dbk();

                            if (!DaoBob.IfExistTier(ExcellArray[BIE_TIERS], folder.Path))
                            {
                                if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                                {
                                    ajoutCompan.Ccustype = "C";
                                    ajoutCompan.Csuptype = "U";
                                }
                                else
                                {
                                    ajoutCompan.Ccustype = "U";
                                    ajoutCompan.Csuptype = "S";
                                }
                                DaoBob.AjoutCompan(ajoutCompan, folder.Path);
                            }
                            else
                            {

                                if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                                {
                                    ajoutCompan.Ccustype = "C";
                                    DaoBob.UpdateCompan(ajoutCompan, "C", folder.Path);
                                }
                                else
                                {
                                    ajoutCompan.Csuptype = "S";
                                    DaoBob.UpdateCompan(ajoutCompan, "S", folder.Path);
                                }
                            }

                        }
                        else
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TYPE TVA TIERS INCORRECT (NA/EX/BE, LU, ...) : " + ExcellArray[BIE_TIERSASS] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                    }


                    // TIERS & INITIALISATION TYPE TVA DU TIERS POUR CONTROLE DU CODE TVA
                    VatType = "N";
                    if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                    {
                        Bob50Tiers customer = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "C", folder.Path);
                        if (customer == null)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> CLIENT INEXISTANT : " + ExcellArray[BIE_TIERS] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            if (customer.Cvatcat != "N")
                            {
                                if (customer.Cvatref != LegisDossier)
                                {
                                    cTemp = customer.Cvatref;
                                    switch (cTemp)
                                    {
                                        case "EX":
                                            VatType = "I";
                                            break;
                                        case "":
                                            VatType = "N";
                                            break;
                                        default:
                                            VatType = "E";
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Bob50Tiers supplier = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "S", folder.Path);
                        if (supplier == null)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> FOURNISSEUR INEXISTANT : " + ExcellArray[BIE_TIERS] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            if (supplier.Cvatcat != "N")
                            {
                                if (supplier.Cvatref != LegisDossier)
                                {
                                    cTemp = supplier.Cvatref;
                                    switch (cTemp)
                                    {
                                        case "EX":
                                            VatType = "I";
                                            break;
                                        case "":
                                            VatType = "N";
                                            break;
                                        default:
                                            VatType = "E";
                                            break;
                                    }
                                }
                            }
                        }
                    }

                    // COMPTE GENERAL
                    DataTable accInfos = DaoBob.getAccountInfos(ExcellArray[BIE_COMPTE], folder.Path);
                    if (accInfos == null)
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COMPTE GENERAL INEXISTANT : " + ExcellArray[BIE_COMPTE] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }
                    else
                    {
                        bool aistitle = Convert.ToBoolean(accInfos.Rows[0]["AISTITLE"]);
                        if (aistitle)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COMPTE GENERAL INVALIDE : " + ExcellArray[BIE_COMPTE] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            // CONTROLE DE L'ANALYTIQUE
                            for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                            {
                                bool accCost = Convert.ToBoolean(accInfos.Rows[0]["COSTR" + ListeSecAna[i]]);
                                if (accCost)
                                {
                                    if (ExcellArray[NbColStd + i].Trim() == "")
                                    {
                                        if (ListeSecAnaMandatory[i] || AnaObligatoireGlobale)
                                        {
                                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> ANALYTIQUE : " + ListeSecAna[i] + " OBLIGATOIRE" + (char)(13) + (char)(10);
                                            FichierOk = false;
                                            LigneOk = false;
                                        }
                                    }
                                    else
                                    {
                                        Tuple<bool, bool> cosecInfos = DaoBob.getCostAnaInfos(ListeSecAna[i], ExcellArray[NbColStd + i], folder.Path);
                                        if (!cosecInfos.Item1)
                                        {
                                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> CODE ANALYTIQUE INVALIDE : " + ExcellArray[NbColStd + i] + (char)(13) + (char)(10);
                                            FichierOk = false;
                                            LigneOk = false;
                                        }
                                        else
                                        {
                                            if (cosecInfos.Item2)
                                            {
                                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> CODE ANALYTIQUE INVALIDE : " + ExcellArray[NbColStd + i] + (char)(13) + (char)(10);
                                                FichierOk = false;
                                                LigneOk = false;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (ExcellArray[NbColStd + i].Trim() != "")
                                    {
                                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COMPTE GENERAL NON ANALYTIQUE : " + ExcellArray[BIE_COMPTE] + " [" + ListeSecAna[i] + "]" + (char)(13) + (char)(10);
                                        FichierOk = false;
                                        LigneOk = false;
                                    }
                                }
                            }
                        }
                    }

                    // MONTANT TVAC
                    if (!Outils.IsNumeric(ExcellArray[BIE_TOTALTTC], "DOUBLE"))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> MONTANT TVAC INVALIDE : " + ExcellArray[BIE_TOTALTTC] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // TVA
                    if (TvaActive && journalInfos.Iswithvat)
                    {
                        ExcellArray[BIE_PRCTVA].Replace("%", "");
                        if (!Outils.IsNumeric(ExcellArray[BIE_PRCTVA], "DOUBLE"))
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TAUX TVA INVALIDE : " + ExcellArray[BIE_PRCTVA] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            VatNat = ExcellArray[BIE_NATTVA];
                            VatPrc = Convert.ToDouble(ExcellArray[BIE_PRCTVA]);
                            tvainfos = DaoBob.GetVATInfos("S", VatType, VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                            if (tvainfos == null)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TVA ERRONEE : " + VatNat + " " + Convert.ToString(VatPrc) + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                            else
                            {
                                if (Outils.Left(journalInfos.Dbtype, 1) != tvainfos.Vsalpur)
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> TVA ERRONEE : " + VatNat + " " + Convert.ToString(VatPrc) + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                        }
                    }

                    // MONTANT BASE
                    if (!Outils.IsNumeric(ExcellArray[BIE_BASE], "DOUBLE"))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> MONTANT BASE INVALIDE : " + ExcellArray[BIE_BASE] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // MONTANT TVA
                    if (!Outils.IsNumeric(ExcellArray[BIE_TVA], "DOUBLE"))
                    {
                        currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> MONTANT TVA INVALIDE : " + ExcellArray[BIE_TVA] + (char)(13) + (char)(10);
                        FichierOk = false;
                        LigneOk = false;
                    }

                    // DEVISE
                    if (ExcellArray[BIE_DEVISE] != "" && ExcellArray[BIE_DEVISE] != DeviseBaseDossier)
                    {
                        Tuple<bool, bool> resultCurr = DaoBob.VerifStdCode("CURRENCY", ExcellArray[BIE_DEVISE], ReadBobIni.CommonDirectory);
                        if (!resultCurr.Item1)
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DEVISE INVALIDE : " + ExcellArray[BIE_DEVISE] + (char)(13) + (char)(10);
                            FichierOk = false;
                            LigneOk = false;
                        }
                        else
                        {
                            if (!resultCurr.Item2)
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> DEVISE INVISIBLE : " + ExcellArray[BIE_DEVISE] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                        }

                        // RECHERCHE D'UN COURS
                        if (!Outils.IsNumeric(ExcellArray[BIE_COURS], "DOUBLE"))
                        {
                            CoursDevise = bob50dao.GetCurrRate(folder, ExcellArray[BIE_DEVISE], DateDoc);

                            if (CoursDevise == (double)(0))
                            {
                                currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COURS INVALIDE : " + ExcellArray[BIE_COURS] + (char)(13) + (char)(10);
                                FichierOk = false;
                                LigneOk = false;
                            }
                            else
                            {
                                if (CoursDevise <= 0)
                                {
                                    currForm.InfosImport.Text = currForm.InfosImport.Text + "LIGNE " + Convert.ToString(NumLigne).PadLeft(3, '0') + " -> COURS INVALIDE : " + ExcellArray[BIE_COURS] + (char)(13) + (char)(10);
                                    FichierOk = false;
                                    LigneOk = false;
                                }
                            }
                        }
                    }
                }

                if (!LigneOk)
                    currForm.InfosImport.Text = currForm.InfosImport.Text + "------------------------------------------------------" + (char)(13) + (char)(10);

                NumLigne = NumLigne + 1;
                //ValCol1Excel = (range.Cells[NumLigne, 1] as Excel.Range).Text;
                //ValCol1Excel.Trim();

                if (currForm.ProgressBar1.Value + 1 < 100)
                    currForm.ProgressBar1.Value = currForm.ProgressBar1.Value + 1;
                else
                    currForm.ProgressBar1.Value = 0;
            }
            #endregion


            // ---------------------------------
            // FICHIER OK LANCEMENT DES LIAISONS 
            // ---------------------------------
            if (FichierOk)
            {
                RuptureDossier = "";
                NumLigne = 2;
                for (int r = 0; r < tableFichierExcel.Rows.Count; r++)
                {

                    // LECTURE LIGNE EXCELL
                    ExcellArray.Clear();
                    for (int i = 0; i <= NbColExcel; i++)
                    {
                        if (i == 13 || i == 14 || i == 15)
                            cTemp = Convert.ToString(Convert.ToDouble(tableFichierExcel.Rows[r][i]));
                        else
                            cTemp = tableFichierExcel.Rows[r][i].ToString();
                        if (cTemp != null)
                            ExcellArray.Add(cTemp.Trim());
                        else
                            ExcellArray.Add("");
                    }

                    if (RuptureDossier != tableFichierExcel.Rows[r][0].ToString())
                    {
                        // OUVERTURE DOSSIER COURANT
                        dossierActuel = tableFichierExcel.Rows[r][0].ToString();
                        RuptureDossier = tableFichierExcel.Rows[r][0].ToString();

                        folder = bob50dao.GetBobFolder(dossierActuel);
                        //conf = bob50dao.GetConfigDossier(folder);
                        conf = DaoBob.LOCAL_GetConfigDossier(folder);

                        DeviseBaseDossier = conf.Basecurrid;
                        LegisDossier = conf.Legis;
                        AnaObligatoireGlobale = conf.Costmand;
                        TvaActive = conf.Prmvatenbld;

                        #region INITIALISATION TABLES DE LIAISON

                        bob50dao.CreateBobLinkFileEntry(folder, "ACH");
                        bob50dao.CreateBobLinkFileEntry(folder, "NCV");
                        bob50dao.CreateBobLinkFileEntry(folder, "NCA");
                        bob50dao.CreateBobLinkFileEntry(folder, "VEN");

                        #endregion
                    }

                    // POSITIONNEMENT SUR LE JOURNAL
                    Bob50Dbk journalInfos = DaoBob.GetJournalInfos(folder, ExcellArray[BIE_JOURNAL], folder.Path);

                    // POSITIONNEMENT SUR LA PERIODE POUR AVOIR L'EXERCICE FISCAL
                    periode = DaoBob.GetPeriode(ExcellArray[BIE_ANNEE], ExcellArray[BIE_MOIS], folder.Path);

                    // POSITIONNEMENT SUR LE CODE TVA BOB
                    if (TvaActive && journalInfos.Iswithvat)
                    {
                        VatType = "N";
                        if (journalInfos.Dbtype == "SAL" || journalInfos.Dbtype == "SAC")
                        {
                            Bob50Tiers customer = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "C", folder.Path);
                            // VarGlobal.BobDossier.Customer.Find(ExcellArray[BIE_TIERS]);
                            if (customer.Cvatcat != "N")
                            {
                                if (customer.Cvatref != LegisDossier)
                                {
                                    cTemp = customer.Cvatref;
                                    switch (cTemp)
                                    {
                                        case "EX":
                                            VatType = "I";
                                            break;
                                        case "":
                                            VatType = "N";
                                            break;
                                        default:
                                            VatType = "E";
                                            break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            Bob50Tiers supplier = DaoBob.getTierInfos(ExcellArray[BIE_TIERS], "S", folder.Path);
                            if (supplier.Cvatcat != "N")
                            {
                                if (supplier.Cvatref != LegisDossier)
                                {
                                    cTemp = supplier.Cvatref;
                                    switch (cTemp)
                                    {
                                        case "EX":
                                            VatType = "I";
                                            break;
                                        case "":
                                            VatType = "N";
                                            break;
                                        default:
                                            VatType = "E";
                                            break;
                                    }
                                }
                            }
                        }
                        VatNat = ExcellArray[BIE_NATTVA].ToUpper();
                        VatPrc = Convert.ToDouble(ExcellArray[BIE_PRCTVA].Replace("%", ""));
                        tvainfos = DaoBob.GetVATInfos("S", VatType, VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                    }

                    // CALCUL DU COURS DE LA DEVISE
                    if (!Outils.IsNumeric(ExcellArray[BIE_COURS], "DOUBLE"))
                    {
                        CoursDevise = (double)(0);

                        if (ExcellArray[BIE_DEVISE].Trim() != "")
                        {
                            CoursDevise = bob50dao.GetCurrRate(folder, ExcellArray[BIE_DEVISE], Convert.ToDateTime(ExcellArray[BIE_DATEDOC]));
                        }
                    }
                    else
                        CoursDevise = Convert.ToDouble(ExcellArray[BIE_COURS]);

                    // GENERATION DE L'ECRITURE
                    // ------------------------
                    if (RuptureDoc != RuptureDossier + ExcellArray[BIE_JOURNAL] + ExcellArray[BIE_NUMDOC])
                    {
                        RuptureDoc = RuptureDossier + ExcellArray[BIE_JOURNAL] + ExcellArray[BIE_NUMDOC];
                        IndImp = (int)(0);
                        TotalBase = (double)(0);
                        TotalTva = (double)(0);
                        TotalTTC = (double)(0);
                        Tva = (double)(0);

                        // ENTETE
                        cTemp = journalInfos.Dbtype;
                        switch (cTemp)
                        {
                            case "SAL":

                                // FT 08/08/2019 -> PAS D'UTILISATION DE AUTONUM EN VENTE CAR BLOQUANT SI NUMEROTATION SANS ANNEE
                                HeaderVEN = new Bob50KhDbk();
                                HeaderVEN.Tdocno = Convert.ToInt32(ExcellArray[BIE_NUMDOC]);
                                HeaderVEN.Tdbk = journalInfos.Dbid;
                                HeaderVEN.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                HeaderVEN.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                HeaderVEN.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                HeaderVEN.TtypeCie = "C";
                                HeaderVEN.TCompan = ExcellArray[BIE_TIERS];
                                NumDocVEN = NumDocVEN + 1;
                                HeaderVEN.Tfyear = periode.Fyear;
                                if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                    HeaderVEN.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEECH]);
                                else
                                    HeaderVEN.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                    HeaderVEN.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                else
                                {
                                    TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                    HeaderVEN.TAmount = TotalTTC;
                                    HeaderVEN.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    HeaderVEN.TCurrency = ExcellArray[BIE_DEVISE];
                                    HeaderVEN.TCurRate = CoursDevise;
                                }
                                HeaderVEN.TRemExt = ExcellArray[BIE_COMM];
                                HeaderVEN.TRemInt = ExcellArray[BIE_REM];
                                HeaderVEN.Tintmode = "S";
                                if (ExcellArray[BIE_NAMEDOC] != "")
                                    HeaderVEN.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);
                                DaoBob.AddEntete(HeaderVEN, folder.Id, "VEN");
                                break;
                            case "SAC":
                                HeaderNCV = new Bob50KhDbk();
                                HeaderNCV.Tdocno = Convert.ToInt32(ExcellArray[BIE_NUMDOC]);
                                HeaderNCV.Tdbk = journalInfos.Dbid;
                                HeaderNCV.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                HeaderNCV.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                HeaderNCV.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                HeaderNCV.TtypeCie = "C";
                                HeaderNCV.TCompan = ExcellArray[BIE_TIERS];
                                HeaderNCV.Tfyear = periode.Fyear;
                                NumDocNCV = NumDocNCV + 1;
                                if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                    HeaderNCV.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                else
                                    HeaderNCV.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                    HeaderNCV.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                else
                                {
                                    TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                    HeaderNCV.TAmount = TotalTTC;
                                    HeaderNCV.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    HeaderNCV.TCurrency = ExcellArray[BIE_DEVISE];
                                    HeaderNCV.TCurRate = CoursDevise;
                                }
                                HeaderNCV.TRemExt = ExcellArray[BIE_COMM];
                                HeaderNCV.TRemInt = ExcellArray[BIE_REM];
                                HeaderNCV.Tintmode = "S";
                                if (ExcellArray[BIE_NAMEDOC] != "")
                                    HeaderNCV.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);

                                DaoBob.AddEntete(HeaderNCV, folder.Id, "NCV");
                                break;
                            case "PUR":
                                NumDocACH = NumDocACH + 1;
                                HeaderACH = new Bob50KhDbk();
                                HeaderACH.Tdocno = Convert.ToInt32(NumDocACH);
                                HeaderACH.Tdbk = journalInfos.Dbid;
                                HeaderACH.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                HeaderACH.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                HeaderACH.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                HeaderACH.TtypeCie = "S";
                                HeaderACH.TCompan = ExcellArray[BIE_TIERS];
                                HeaderACH.Tfyear = periode.Fyear;
                                if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                    HeaderACH.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEECH]);
                                else
                                    HeaderACH.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                    HeaderACH.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                else
                                {
                                    TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                    HeaderACH.TAmount = TotalTTC;
                                    HeaderACH.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    HeaderACH.TCurrency = ExcellArray[BIE_DEVISE];
                                    HeaderACH.TCurRate = CoursDevise;
                                }
                                HeaderACH.TRemExt = ExcellArray[BIE_COMM];
                                HeaderACH.TRemInt = ExcellArray[BIE_REM];
                                HeaderACH.Tintmode = "S";
                                DaoBob.AddEntete(HeaderACH, folder.Id, "ACH");
                                if (ExcellArray[BIE_NAMEDOC] != "")
                                    HeaderACH.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);

                                break;
                            case "PUC":
                                NumDocNCA = NumDocNCA + 1;
                                HeaderNCA = new Bob50KhDbk();
                                HeaderNCA.Tdocno = Convert.ToInt32(NumDocNCA);
                                HeaderNCA.Tdbk = journalInfos.Dbid;
                                HeaderNCA.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                                HeaderNCA.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                                HeaderNCA.TdocDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                HeaderNCA.TtypeCie = "S";
                                HeaderNCA.TCompan = ExcellArray[BIE_TIERS];
                                HeaderNCA.Tfyear = periode.Fyear;
                                if (Outils.IsDate(ExcellArray[BIE_DATEECH]))
                                    HeaderNCA.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                else
                                    HeaderNCA.TdueDate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                                if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                                    HeaderNCA.TAmount = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                else
                                {
                                    TotalTTC = Math.Round(Convert.ToDouble(ExcellArray[BIE_TOTALTTC]) / CoursDevise, 2);
                                    HeaderNCA.TAmount = TotalTTC;
                                    HeaderNCA.TCurAmn = Convert.ToDouble(ExcellArray[BIE_TOTALTTC]);
                                    HeaderNCA.TCurrency = ExcellArray[BIE_DEVISE];
                                    HeaderNCA.TCurRate = CoursDevise;
                                }
                                HeaderNCA.TRemExt = ExcellArray[BIE_COMM];
                                HeaderNCA.TRemInt = ExcellArray[BIE_REM];
                                HeaderNCA.Tintmode = "S";
                                if (ExcellArray[BIE_NAMEDOC] != "")
                                    HeaderNCA.TPdfFileName = Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], ExcellArray[BIE_NAMEDOC]);

                                DaoBob.AddEntete(HeaderNCA, folder.Id, "NCA");
                                break;
                        }
                    }

                    // IMPUTATION
                    IndImp = IndImp + (int)(1);
                    cTemp = journalInfos.Dbtype;
                    tvainfos = DaoBob.GetVATInfos("S", VatType, VatNat, Convert.ToString(VatPrc).Replace(',', '.'), folder.Path);
                    switch (cTemp)
                    {
                        case "SAL":
                            LineVEN = new Bob50KlDbk();
                            LineVEN.Tdocno = NumDocVEN;
                            LineVEN.Tdbk = journalInfos.Dbid;
                            LineVEN.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                            LineVEN.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                            LineVEN.Tdocline = IndImp;
                            LineVEN.Ttypeline = "S";
                            LineVEN.Tacttype = "";
                            LineVEN.Taccount = ExcellArray[BIE_COMPTE];
                            LineVEN.Tfyear = periode.Fyear;
                            LineVEN.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                            if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                            {
                                BaseTva = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                Tva = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                TotalBase = TotalBase + BaseTva;
                                TotalTva = TotalTva + Tva;
                                LineVEN.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineVEN.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineVEN.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineVEN.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineVEN.TVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineVEN.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            else
                            {
                                BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                TotalBase = TotalBase + BaseTva;
                                TotalTva = TotalTva + Tva;
                                LineVEN.Tamount = BaseTva;
                                LineVEN.TBasVat = BaseTva;
                                LineVEN.TBasLstAmn = BaseTva;
                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineVEN.TVatTotAmn = Tva;
                                    TvaDeduc = Math.Round(Tva / 100 * tvainfos.Vdeduc, 2);
                                    LineVEN.TVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineVEN.TVcDblAmn = Tva;
                                }

                                if (TvaActive && journalInfos.Iswithvat)
                                {

                                    LineVEN.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineVEN.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineVEN.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineVEN.TCurVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineVEN.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            if (TvaActive && journalInfos.Iswithvat)
                                LineVEN.TVStored = tvainfos.Vstored;
                            LineVEN.Tdc = "C";
                            LineVEN.TRem = ExcellArray[BIE_REM];

                            // ANALYTIQUE
                            for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                            {
                                switch (i)
                                {
                                    case 0:
                                        LineVEN.Cost_1 = ExcellArray[NbColStd];
                                        break;
                                    case 1:
                                        LineVEN.Cost_2 = ExcellArray[NbColStd + 1];
                                        break;
                                    case 2:
                                        LineVEN.Cost_3 = ExcellArray[NbColStd + 2];
                                        break;
                                    case 3:
                                        LineVEN.Cost_4 = ExcellArray[NbColStd + 3];
                                        break;
                                    case 4:
                                        LineVEN.Cost_5 = ExcellArray[NbColStd + 4];
                                        break;
                                    case 5:
                                        LineVEN.Cost_6 = ExcellArray[NbColStd + 5];
                                        break;
                                    case 6:
                                        LineVEN.Cost_7 = ExcellArray[NbColStd + 6];
                                        break;
                                    case 7:
                                        LineVEN.Cost_8 = ExcellArray[NbColStd + 7];
                                        break;
                                }
                            }

                            DaoBob.AddLigne(LineVEN, ListeSecAna, folder.Id, "VEN");
                            break;
                        case "SAC":
                            LineNCV = new Bob50KlDbk();
                            LineNCV.Tdocno = NumDocNCV;
                            LineNCV.Tdbk = journalInfos.Dbid;
                            LineNCV.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                            LineNCV.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                            LineNCV.Tdocline = IndImp;
                            LineNCV.Ttypeline = "S";
                            LineNCV.Tacttype = "";
                            LineNCV.Taccount = ExcellArray[BIE_COMPTE];
                            LineNCV.Tfyear = periode.Fyear;
                            LineNCV.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                            if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                            {
                                LineNCV.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineNCV.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineNCV.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineNCV.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineNCV.TVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineNCV.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            else
                            {
                                BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                TotalBase = TotalBase + BaseTva;
                                TotalTva = TotalTva + Tva;
                                LineNCV.Tamount = BaseTva;
                                LineNCV.TBasVat = BaseTva;
                                LineNCV.TBasLstAmn = BaseTva;

                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineNCV.TVatTotAmn = Tva;
                                    LineNCV.TVatAmn = Tva;
                                    if (tvainfos.Vkind == "D")
                                        LineNCV.TVatDblAmn = Tva;
                                }

                                LineNCV.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineNCV.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);

                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineNCV.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineNCV.TCurVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineNCV.TVcDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            if (TvaActive && journalInfos.Iswithvat)
                                LineNCV.TVStored = tvainfos.Vstored;
                            LineNCV.Tdc = "D";
                            LineNCV.TRem = ExcellArray[BIE_REM];

                            // ANALYTIQUE
                            for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                            {
                                switch (i)
                                {
                                    case 0:
                                        LineNCV.Cost_1 = ExcellArray[NbColStd];
                                        break;
                                    case 1:
                                        LineNCV.Cost_2 = ExcellArray[NbColStd + 1];
                                        break;
                                    case 2:
                                        LineNCV.Cost_3 = ExcellArray[NbColStd + 2];
                                        break;
                                    case 3:
                                        LineNCV.Cost_4 = ExcellArray[NbColStd + 3];
                                        break;
                                    case 4:
                                        LineNCV.Cost_5 = ExcellArray[NbColStd + 4];
                                        break;
                                    case 5:
                                        LineNCV.Cost_6 = ExcellArray[NbColStd + 5];
                                        break;
                                    case 6:
                                        LineNCV.Cost_7 = ExcellArray[NbColStd + 6];
                                        break;
                                    case 7:
                                        LineNCV.Cost_8 = ExcellArray[NbColStd + 7];
                                        break;
                                }
                            }

                            DaoBob.AddLigne(LineNCV, ListeSecAna, folder.Id, "NCV");
                            break;
                        case "PUR":
                            LineACH = new Bob50KlDbk();
                            LineACH.Tdocno = NumDocACH;
                            LineACH.Tdbk = journalInfos.Dbid;
                            LineACH.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                            LineACH.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                            LineACH.Tdocline = IndImp;
                            LineACH.Ttypeline = "S";
                            LineACH.Tacttype = "";
                            LineACH.Taccount = ExcellArray[BIE_COMPTE];
                            LineACH.Tfyear = periode.Fyear;
                            LineACH.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                            if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                            {
                                LineACH.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineACH.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineACH.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineACH.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineACH.TVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineACH.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            else
                            {
                                BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                TotalBase = TotalBase + BaseTva;
                                TotalTva = TotalTva + Tva;
                                LineACH.Tamount = BaseTva;
                                LineACH.TBasVat = BaseTva;
                                LineACH.TBasLstAmn = BaseTva;

                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineACH.TVatTotAmn = Tva;
                                    LineACH.TVatAmn = Tva;
                                    if (tvainfos.Vkind == "D")
                                        LineACH.TVatDblAmn = Tva;
                                }

                                LineACH.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineACH.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);

                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineACH.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineACH.TCurVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineACH.TVcDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            if (TvaActive && journalInfos.Iswithvat)
                                LineACH.TVStored = tvainfos.Vstored;
                            LineACH.Tdc = "D";
                            LineACH.TRem = ExcellArray[BIE_REM];

                            // ANALYTIQUE
                            for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                            {
                                switch (i)
                                {
                                    case 0:
                                        LineACH.Cost_1 = ExcellArray[NbColStd];
                                        break;
                                    case 1:
                                        LineACH.Cost_2 = ExcellArray[NbColStd + 1];
                                        break;
                                    case 2:
                                        LineACH.Cost_3 = ExcellArray[NbColStd + 2];
                                        break;
                                    case 3:
                                        LineACH.Cost_4 = ExcellArray[NbColStd + 3];
                                        break;
                                    case 4:
                                        LineACH.Cost_5 = ExcellArray[NbColStd + 4];
                                        break;
                                    case 5:
                                        LineACH.Cost_6 = ExcellArray[NbColStd + 5];
                                        break;
                                    case 6:
                                        LineACH.Cost_7 = ExcellArray[NbColStd + 6];
                                        break;
                                    case 7:
                                        LineACH.Cost_8 = ExcellArray[NbColStd + 7];
                                        break;
                                }
                            }

                            DaoBob.AddLigne(LineACH, ListeSecAna, folder.Id, "ACH");
                            break;
                        case "PUC":
                            LineNCA = new Bob50KlDbk();
                            LineNCA.Tdocno = NumDocNCA;
                            LineNCA.Tdbk = journalInfos.Dbid;
                            LineNCA.Tyear = Convert.ToInt16(ExcellArray[BIE_ANNEE]);
                            LineNCA.Tmonth = Convert.ToInt16(ExcellArray[BIE_MOIS]);
                            LineNCA.Tdocline = IndImp;
                            LineNCA.Ttypeline = "S";
                            LineNCA.Tacttype = "";
                            LineNCA.Taccount = ExcellArray[BIE_COMPTE];
                            LineNCA.Tfyear = periode.Fyear;
                            LineNCA.Tdocdate = Convert.ToDateTime(ExcellArray[BIE_DATEDOC]);
                            if (ExcellArray[BIE_DEVISE] == "" || ExcellArray[BIE_DEVISE] == DeviseBaseDossier)
                            {
                                LineNCA.Tamount = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineNCA.TBasVat = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineNCA.TBasLstAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                    LineNCA.TVatTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineNCA.TVatAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineNCA.TVatDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            else
                            {
                                BaseTva = Math.Round(Convert.ToDouble(ExcellArray[BIE_BASE]) / CoursDevise, 2);
                                Tva = Math.Round(Convert.ToDouble(ExcellArray[BIE_TVA]) / CoursDevise, 2);
                                TotalBase = TotalBase + BaseTva;
                                TotalTva = TotalTva + Tva;
                                LineNCA.Tamount = BaseTva;
                                LineNCA.TBasVat = BaseTva;
                                LineNCA.TBasLstAmn = BaseTva;

                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineNCA.TVatTotAmn = Tva;
                                    LineNCA.TVatAmn = Tva;
                                    if (tvainfos.Vkind == "D")
                                        LineNCA.TVatDblAmn = Tva;
                                }

                                LineNCA.TCurAmn = Convert.ToDouble(ExcellArray[BIE_BASE]);
                                LineNCA.TCbVat = Convert.ToDouble(ExcellArray[BIE_BASE]);

                                if (TvaActive && journalInfos.Iswithvat)
                                {
                                    LineNCA.TVcTotAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                    TvaDeduc = Math.Round((Convert.ToDouble(ExcellArray[BIE_TVA]) / 100) * tvainfos.Vdeduc, 2);
                                    LineNCA.TCurAmn = TvaDeduc;
                                    if (tvainfos.Vkind == "D")
                                        LineNCA.TVcDblAmn = Convert.ToDouble(ExcellArray[BIE_TVA]);
                                }
                            }
                            if (TvaActive && journalInfos.Iswithvat)
                                LineNCA.TVStored = tvainfos.Vstored;
                            LineNCA.Tdc = "C";
                            LineNCA.TRem = ExcellArray[BIE_REM];

                            // ANALYTIQUE
                            for (int i = 0; i <= ListeSecAna.Count - 1; i++)
                            {
                                switch (i)
                                {
                                    case 0:
                                        LineNCA.Cost_1 = ExcellArray[NbColStd];
                                        break;
                                    case 1:
                                        LineNCA.Cost_2 = ExcellArray[NbColStd + 1];
                                        break;
                                    case 2:
                                        LineNCA.Cost_3 = ExcellArray[NbColStd + 2];
                                        break;
                                    case 3:
                                        LineNCA.Cost_4 = ExcellArray[NbColStd + 3];
                                        break;
                                    case 4:
                                        LineNCA.Cost_5 = ExcellArray[NbColStd + 4];
                                        break;
                                    case 5:
                                        LineNCA.Cost_6 = ExcellArray[NbColStd + 5];
                                        break;
                                    case 6:
                                        LineNCA.Cost_7 = ExcellArray[NbColStd + 6];
                                        break;
                                    case 7:
                                        LineNCA.Cost_8 = ExcellArray[NbColStd + 7];
                                        break;
                                }
                            }

                            DaoBob.AddLigne(LineNCA, ListeSecAna, folder.Id, "NCA");
                            break;
                    }

                    bool rupture = false;
                    if (r + 1 < tableFichierExcel.Rows.Count)
                    {
                        cTemp = tableFichierExcel.Rows[r + 1][0].ToString().Trim() + tableFichierExcel.Rows[r + 1][5].ToString().Trim();
                        if (RuptureDoc != cTemp)
                            rupture = true;
                    }
                    else
                    {
                        rupture = true;
                    }

                    //// CONTROLES DE FIN DE RUPTURE
                    if (rupture)
                    {
                        if (ExcellArray[BIE_DEVISE] != "" && ExcellArray[BIE_DEVISE] != DeviseBaseDossier)
                        {
                            Ecart = Math.Round(TotalTTC - Math.Round(TotalBase + TotalTva, 2), 2);
                            if (Ecart != (double)(0))
                            {
                                if (Tva == (double)(0))
                                {
                                    cTemp = journalInfos.Dbtype;
                                    switch (cTemp)
                                    {
                                        case "SAL":
                                            BaseTva = Math.Round(Convert.ToDouble(LineVEN.Tamount) + Ecart, 2);
                                            LineVEN.Tamount = BaseTva;
                                            LineVEN.TBasVat = BaseTva;
                                            LineVEN.TBasLstAmn = BaseTva;
                                            DaoBob.UpdateLigneBASE(LineVEN, folder.Id, "VEN");
                                            break;
                                        case "SAC":
                                            BaseTva = Math.Round(Convert.ToDouble(LineNCV.Tamount) + Ecart, 2);
                                            LineNCV.Tamount = BaseTva;
                                            LineNCV.TBasVat = BaseTva;
                                            LineNCV.TBasLstAmn = BaseTva;
                                            DaoBob.UpdateLigneBASE(LineNCV, folder.Id, "NCV");
                                            break;
                                        case "PUR":
                                            BaseTva = Math.Round(Convert.ToDouble(LineACH.Tamount) + Ecart, 2);
                                            LineACH.Tamount = BaseTva;
                                            LineACH.TBasVat = BaseTva;
                                            LineACH.TBasLstAmn = BaseTva;
                                            DaoBob.UpdateLigneBASE(LineACH, folder.Id, "ACH");
                                            break;
                                        case "PUC":
                                            BaseTva = Math.Round(Convert.ToDouble(LineNCA.Tamount) + Ecart, 2);
                                            LineNCA.Tamount = BaseTva;
                                            LineNCA.TBasVat = BaseTva;
                                            LineNCA.TBasLstAmn = BaseTva;
                                            DaoBob.UpdateLigneBASE(LineNCA, folder.Id, "NCA");
                                            break;
                                    }
                                }
                                else
                                {
                                    if (TvaActive && journalInfos.Iswithvat)
                                    {
                                        cTemp = journalInfos.Dbtype;
                                        switch (cTemp)
                                        {
                                            case "SAL":
                                                Tva = Math.Round(Convert.ToDouble(LineVEN.TVatAmn) + Ecart, 2);
                                                LineVEN.TVatAmn = Tva;
                                                LineVEN.TVatTotAmn = Tva;
                                                DaoBob.UpdateLigneTVA(LineVEN, folder.Id, "VEN");
                                                break;
                                            case "SAC":
                                                Tva = Math.Round(Convert.ToDouble(LineNCV.TVatAmn) + Ecart, 2);
                                                LineNCV.TVatAmn = Tva;
                                                LineNCV.TVatTotAmn = Tva;
                                                DaoBob.UpdateLigneTVA(LineNCV, folder.Id, "NCV");
                                                break;
                                            case "PUR":
                                                Tva = Math.Round(Convert.ToDouble(LineACH.TVatAmn) + Ecart, 2);
                                                LineACH.TVatAmn = Tva;
                                                LineACH.TVatTotAmn = Tva;
                                                DaoBob.UpdateLigneTVA(LineACH, folder.Id, "ACH");
                                                break;
                                            case "PUC":
                                                Tva = Math.Round(Convert.ToDouble(LineNCA.TVatAmn) + Ecart, 2);
                                                LineNCA.TVatAmn = Tva;
                                                LineNCA.TVatTotAmn = Tva;
                                                DaoBob.UpdateLigneTVA(LineNCA, folder.Id, "NCA");
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (currForm.ProgressBar1.Value + 1 < 100)
                    currForm.ProgressBar1.Value = currForm.ProgressBar1.Value + 1;
                else
                    currForm.ProgressBar1.Value = 0;
            }

            // REINITIALISATION PROGRESS BAR
            currForm.ProgressBar1.Value = (int)(1);

            // FICHIER EN ERREUR
            if (!FichierOk)
            {
                MessageBox.Show("Erreurs dans le fichier d'import, veuillez corriger", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // LIAISONS VERS LES DIFFERENTS DOSSIERS EN MODE TEST
            LinkOk = true;
            for (int x = 0; x < ListeDossiers.Count; x++)
            {

                // LIAISON DES VENTES
                if (NumDocVEN > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHVEN.ADT /LNF=LINK_KLVEN.ADT /NOPOST /MOD=ENTRYSAL /SEPARATE /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG")))
                    {
                        currForm.InfosImport.Text = "DOSSIER " + ListeDossiers[x] + " ERREUR LIAISON DES VENTES" + (char)(13) + (char)(10);
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                // LIAISON DES NOTES DE CREDIT
                if (NumDocNCV > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDFLINK_K=HNCV.ADT /LNF=LINK_KLNCV.ADT /SEPARATE /NOPOST /MOD=ENTRYSAL /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG")))
                    {
                        currForm.InfosImport.Text = "DOSSIER " + ListeDossiers[x] + " ERREUR LIAISON DES NOTES DE CREDIT VENTES" + (char)(13) + (char)(10);
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                // LIAISON DES ACHATS
                if (NumDocACH > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHACH.ADT /LNF=LINK_KLACH.ADT /SEPARATE /NOPOST /MOD=ENTRYPUR /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG")))
                    {
                        currForm.InfosImport.Text = "DOSSIER " + ListeDossiers[x] + " ERREUR LIAISON DES ACHATS" + (char)(13) + (char)(10);
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }

                // LIAISON DES !ES DE CREDIT SUR ACHAT
                if (NumDocNCA > (long)(0))
                {
                    RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHNCA.ADT /LNF=LINK_KLNCA.ADT /SEPARATE /NOPOST /MOD=ENTRYPUR /AUTOCLOSE");
                    if (File.Exists(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG")))
                    {
                        currForm.InfosImport.Text = "DOSSIER " + ListeDossiers[x] + " ERREUR LIAISON DES NOTES DE CREDIT ACHATS" + (char)(13) + (char)(10);
                        Dev4LogLayerFiles.TxtReader BobLogReader = new Dev4LogLayerFiles.TxtReader();
                        BobLogReader.openStreamReader(Path.Combine(VarGlobal.BOBLinkPath, ListeDossiers[x], "LNK.LOG"), Encoding.GetEncoding(1252));
                        cTemp = BobLogReader.readLine();
                        while (cTemp != "eof")
                        {
                            currForm.InfosImport.Text = currForm.InfosImport.Text + cTemp + (char)(13) + (char)(10);
                            cTemp = BobLogReader.readLine();
                        }
                        BobLogReader.closeStreamReader();
                        LinkOk = false;
                    }
                }
            }


            // LIAISON REELLE POUR TOUS LES DOSSIERS
            if (LinkOk)
            {

                for (int x = 0; x < ListeDossiers.Count; x++)
                {
                    VarGlobal.listPdf.Clear();

                    // LIAISON DES VENTES
                    if (NumDocVEN > (long)(0))
                    {
                        RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHVEN.ADT /LNF=LINK_KLVEN.ADT /INTEMP /MOD=ENTRYSAL /AUTOCLOSE");
                        
                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "VEN");
                    }

                    // LIAISON DES NOTES DE CREDIT
                    if (NumDocNCV > (long)(0))
                    {
                        RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHNCV.ADT /LNF=LINK_KLNCV.ADT /INTEMP /MOD=ENTRYSAL /AUTOCLOSE");
                        
                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "NCV");
                    }

                    // LIAISON DES ACHATS
                    if (NumDocACH > (long)(0))
                    {
                        RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHACH.ADT /LNF=LINK_KLACH.ADT /INTEMP /MOD=ENTRYPUR /AUTONUM /AUTOCLOSE");

                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "ACH");
                    }

                    // LIAISON DES !ES DE CREDIT SUR ACHAT
                    if (NumDocNCA > (long)(0))
                    {
                        RetVal = Outils.ExecCmd(AppSettingsManager.Dictionnary["BOB_PATH"], "BOBLINKADSADT.EXE", "/USR=" + VarGlobal.BOBUser + " /CIE=" + ListeDossiers[x] + " /HDF=LINK_KHNCA.ADT /LNF=LINK_KLNCA.ADT /INTEMP /MOD=ENTRYPUR /AUTONUM /AUTOCLOSE");

                        DaoBob.GetListPdfImported(VarGlobal.BOBDossierRun, "NCA");
                    }

                    // ARCHIVAGE DES PDF
                    if (!Directory.Exists(Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], "Archives")))
                        Directory.CreateDirectory(Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], "Archives"));
                    foreach (string pdf in VarGlobal.listPdf)
                    {
                        if (File.Exists(pdf))
                        {
                            File.Move(pdf, Path.Combine(AppSettingsManager.Dictionnary["PATH_DOCUMENTS"], "Archives", Path.GetFileName(pdf)));
                        }
                    }
                }

                currForm.FichierImport.Text = "";
                MessageBox.Show("Import terminé", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Problèmes lors de l'import, contacter votre revendeur", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// LANCEMENT DE LA FENETRE DE GESTION DE LA LICENCE
        /// </summary>
        public void GestionLicence()
        {
            String client = dev4LogBob50dao.GetBobLicence("NAME");

            ControllerLicense ctrlLicence = new ControllerLicense(client, "");
        }

        /// <summary>
        /// LANCEMENT DE LA FENETRE DE GESTION DES LOGS
        /// </summary>
        public void GestionDesLogs()
        {
            String client = dev4LogBob50dao.GetBobLicence("NAME");

            ControllerLogs ctrlLogs = new ControllerLogs(client);
        }


        /// <summary>
        /// LANCEMENT DE LA FENETRE DE CONFIGURATION MAIL DEV4LOG AZUR
        /// </summary>
        public void SettingsMailDev4LogAzur()
        {
            ControllerSettingsMail ctrlSettingsMail = new ControllerSettingsMail(true);
        }
    }
}
