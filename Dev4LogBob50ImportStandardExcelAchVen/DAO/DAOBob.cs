using System;
using System.Collections.Generic;
using System.Data;
using Dev4LogLayer.Daos;
using Dev4LogBob50ImportStandardExcelAchVen.Models;
using Dev4LogLayer.Bob50.Models;
using Dev4LogLayer.Bob50;
using Dev4LogLayer.Globals;
using System.Windows.Forms;
using Dev4LogLayer.Exceptions;
using System.Text;

namespace Dev4LogBob50ImportStandardExcelAchVen.DAO
{
    public static class DaoBob
    {
        static readonly DaoADS daoADS = new DaoADS();

        public static Boolean GetJournalOk(String codeJournal, string dataPath)
        {
            // CONTROLE EXISTANCE JOURNAL
            Boolean ok = false;
            DataTable table;
            try
            {
                daoADS.Directory = dataPath;

                String query = $@"SELECT * FROM AC_DBK WHERE DBID = '{codeJournal}' ORDER BY DBID";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    ok = true;
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : GetJournalOk() in DaoBob.cs.\n" + ex.Message);
            }
            return ok;
        }

        public static Bob50Dbk GetJournalInfos(Bob50Folder folder, string code, string dataPath)
        {
            Bob50Dbk bob50Dbk = new Bob50Dbk();

            try
            {
                daoADS.Directory = dataPath;
                DataTable dataTable = daoADS.ExecuteQuery("SELECT * FROM AC_DBK WHERE dbid = '" + code + "'");
                bob50Dbk.Dbid = code;
                bob50Dbk.Dbaccount = (DBNull.Value.Equals(dataTable.Rows[0]["dbaccount"]) ? "" : dataTable.Rows[0]["dbaccount"].ToString());
                bob50Dbk.Dbendamn = (DBNull.Value.Equals(dataTable.Rows[0]["dbendamn"]) ? 0.0 : Convert.ToDouble(dataTable.Rows[0]["dbendamn"]));
                bob50Dbk.Dbtype = dataTable.Rows[0]["dbtype"].ToString();
                bob50Dbk.Heading1 = dataTable.Rows[0]["heading1"].ToString();
                bob50Dbk.Iswithvat = !DBNull.Value.Equals(dataTable.Rows[0]["ISWITHVAT"]) && Convert.ToBoolean(dataTable.Rows[0]["ISWITHVAT"]);
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : GetJournalOk() in DaoBob.cs.\n" + ex.Message);
            }
            return bob50Dbk;
        }

        public static Boolean IfExistTier(String cid, string dataPath)
        {
            // CONTROLE EXISTANCE JOURNAL
            Boolean ok = false;
            DataTable table;
            try
            {
                daoADS.Directory = dataPath;

                String query = $@"SELECT * FROM AC_COMPAN WHERE UPPER(CID) = '{cid.ToUpper().Replace("'", "''")}'";
                Log.WriteLog("TraceDebug", "IfExistTier - Query " + query);
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    ok = true;
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : IfExistTier() in DaoBob.cs.\n" + ex.Message);
            }
            return ok;
        }

        public static int IfExistCompBk(String cid, string dataPath)
        {
            // CONTROLE EXISTANCE COMPBK
            int result = 0;
            DataTable table;
            try
            {

                daoADS.Directory = dataPath;

                String query = $@"SELECT * FROM AC_COMPBK WHERE CID = '{cid.Replace("'", "''")}'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count >= 1)
                {
                    result = table.Rows.Count;
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : IfExistCompBk() in DaoBob.cs.\n" + ex.Message);
            }
            return result;
        }

        public static bool IfCompBkIbanExist(String cid, String banknoindex, string dataPath)
        {
            // CONTROLE EXISTANCE COMPBK
            bool result = false;
            DataTable table;
            try
            {

                daoADS.Directory = dataPath;

                String query = $@"SELECT * FROM AC_COMPBK WHERE CID = '{cid.Replace("'", "''")}' AND BANKNOINDEX = '{banknoindex}'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count >= 1)
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : IfCompBkIbanExist() in DaoBob.cs.\n" + ex.Message);
            }
            return result;
        }

        public static bool IfExistCompDe(String cid, string dataPath)
        {
            // CONTROLE EXISTANCE JOURNAL
            bool result = false;
            DataTable table;
            try
            {
                daoADS.Directory = dataPath;

                String query = $@"SELECT * FROM AC_COMPDE WHERE CID = '{cid.Replace("'", "''")}'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : IfExistCompDe() in DaoBob.cs.\n" + ex.Message);
            }
            return result;
        }

        public static Bob50Tiers getTierInfos(string id, string type = "C", string dataPath = "")
        {
            Bob50Tiers tier = null;

            DataTable table;
            try
            {
                daoADS.Directory = dataPath;

                String query;
                if (type == "C")
                    query = $@"SELECT * FROM AC_COMPAN WHERE UPPER(CID) = '{id.ToUpper().Replace("'", "''")}' AND CCUSTYPE = 'C'";
                else
                    query = $@"SELECT * FROM AC_COMPAN WHERE UPPER(CID) = '{id.ToUpper().Replace("'", "''")}' AND CSUPTYPE = 'S'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    tier = new Bob50Tiers();
                    tier.Cvatcat = table.Rows[0]["CVATCAT"].ToString();
                    tier.Cvatref = table.Rows[0]["CVATREF"].ToString();
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : getTierInfos() in DaoBob.cs.\n" + ex.Message);
            }
            return tier;
        }

        public static DataTable getAccountInfos(string aid, string dataPath)
        {
            DataTable table;
            DataTable result = null;
            try
            {
                daoADS.Directory = dataPath;

                String query;
                query = $@"SELECT * FROM AC_ACCOUN WHERE AID = '{aid}'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    result = table;
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : getAccountInfos() in DaoBob.cs.\n" + ex.Message);
            }

            return result;
        }

        public static Tuple<bool, bool> getCostAnaInfos(string section, string code, String dataPath)
        {
            bool result = false;
            bool aistitle = false;

            DataTable table;
            try
            {
                daoADS.Directory = dataPath;

                String query;
                query = $@"SELECT * FROM AC_COSECT WHERE ASECTION = '{section}' AND AID = '{code}'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    result = true;
                    aistitle = DBNull.Value.Equals(table.Rows[0]["AISTITLE"]) ? false : Convert.ToBoolean(table.Rows[0]["AISTITLE"]);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : getCostAnaInfos() in DaoBob.cs.\n" + ex.Message);
            }

            return new Tuple<bool, bool>(result, aistitle);
        }

        public static Tuple<bool, bool> VerifStdCode(String name, String code, String dossier)
        {
            bool ok = false;
            DataTable table;
            bool visible = false;
            try
            {
                daoADS.Directory = dossier;
                String query = $@"SELECT * FROM STDCODE WHERE TNAME = '{name}' AND TCODE='{code}'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    ok = true;
                    visible = Convert.ToBoolean(table.Rows[0]["VISIBLE"]);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : VerifStdCode() in DaoBob.cs.\n" + ex.Message);
            }
            return new Tuple<bool, bool>(ok, visible);
        }

        public static Bob50Period GetPeriode(String année, String mois, String dataPath)
        {

            // RENVOI UNE PERIODE
            Bob50Period periode = new Bob50Period();
            periode.Year = 0;
            periode.Month = 0;
            DataTable table;
            try
            {
                daoADS.Directory = dataPath;
                String query = "SELECT * FROM AC_PERIOD WHERE YEAR = " + année + " AND MONTH = " + mois + " AND ISNULL(PURCLOSED, FALSE) = FALSE";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    periode = new Bob50Period(); 
                    periode.Label = table.Rows[0]["LABEL"].ToString();
                    periode.Fyear = table.Rows[0]["FYEAR"].ToString();
                    periode.Year = Convert.ToInt32(table.Rows[0]["YEAR"]);
                    periode.Month = Convert.ToInt32(table.Rows[0]["MONTH"]);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : GetPeriode() in DaoBob.cs.\n" + ex.Message);
            }

            return periode;
        }

        public static String GetSecAnaObligatoire(String codeSection, String dataPath)
        {

            // RENVOI SI SECTION OBLIGATOIRE
            String obligatoire = "";
            DataTable table;
            try
            {
                daoADS.Directory = dataPath;
                String query = "SELECT ISREQUIRED FROM AC_CODEF WHERE CODE = '" + codeSection + "'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    obligatoire = table.Rows[0]["ISREQUIRED"].ToString().ToUpper();
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : GetSecAnaObligatoire() in DaoBob.cs.\n" + ex.Message);
            }
            return obligatoire;
        }

        public static String GetTypeSecAna(String codeSection, String dataPath)
        {

            // RENVOI SI SECTION OBLIGATOIRE
            String obligatoire = "";
            DataTable table;
            try
            {
                daoADS.Directory = dataPath;
                String query = "SELECT ISNULL(PTYPE, '') AS PTYPE FROM AC_CODEF WHERE CODE = '" + codeSection + "'";
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    obligatoire = table.Rows[0]["PTYPE"].ToString().ToUpper();
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : GetSecAnaObligatoire() in DaoBob.cs.\n" + ex.Message);
            }
            return obligatoire;
        }

        public static Boolean AddEntete(Bob50KhDbk Entete, String dossier, String type = "ACH")
        {
            String reqSql = "";
            Boolean ok;


            // CREATION DES ENTETES
            try
            {
                daoADS.Directory = ReadBobIni.LinkDirectory + @"\" + dossier;
                reqSql = $"INSERT INTO LINK_KH{type} (TDBK, TFYEAR, TYEAR, TMONTH, TDOCNO, TINTMODE, TDOCDATE, TTYPCIE, TCOMPAN, TDUEDATE, TAMOUNT, TCURAMN, TCURRENCY, TCURRATE, TREMEXT, TREMINT, TPDFFILENAME) VALUES ('" +
                                                                Entete.Tdbk + "', '" +
                                                                Entete.Tfyear + "', " +
                                                                Entete.Tyear + ", " +
                                                                Entete.Tmonth + ", " +
                                                                Entete.Tdocno + ", '" +
                                                                Entete.Tintmode + "', '" +
                                                                Entete.TdocDate.ToString("yyyy-MM-dd") + "', '" +
                                                                Entete.TtypeCie + "', '" +
                                                                Entete.TCompan.Replace("'", "''") + "', '" +
                                                                Entete.TdueDate.ToString("yyyy-MM-dd") + "', " +
                                                                Entete.TAmount.ToString().Replace(",", ".") + ", " +
                                                                Entete.TCurAmn.ToString().Replace(",", ".") + ", '" + 
                                                                Entete.TCurrency + "', " +
                                                                Entete.TCurRate.ToString().Replace(",", ".") + ", '" +
                                                                Entete.TRemExt.Replace("'", "''") + "', '" +
                                                                Entete.TRemInt.Replace("'", "''") + "', '" +
                                                                Entete.TPdfFileName + "')";
                daoADS.ExecuteQuery(reqSql);
                ok = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : AddEntete() in DaoBob.cs.\n" + reqSql + "\n" + ex.Message);
                ok = false;
            }
            return ok;
        }

        public static Boolean UpdateLigneBASE(Bob50KlDbk Ligne, String dossier, String type = "ACH")
        {
            Boolean ok;
            string reqSql;
            try
            {
                ok = true;
                daoADS.Directory = ReadBobIni.LinkDirectory + @"\" + dossier;
                reqSql = $"UPDATE LINK_KL{type} SET TAMOUNT = " + Ligne.Tamount.ToString().Replace(",", ".") + ", " + 
                        "TBASVAT = " + Ligne.TBasVat.ToString().Replace(",", ".") + ", " + 
                        "TBASLSTAMN = " + Ligne.TBasLstAmn.ToString().Replace(",", ".") + " " +
                        "WHERE TDBK = '" + Ligne.Tdbk + "' AND TFYEAR = '" + Ligne.Tfyear + "' AND TMONTH = " + Ligne.Tmonth + " AND " +
                              "TDOCNO = " + Ligne.Tdocno + " AND TDOCLINE = " + Ligne.Tdocline;
                daoADS.ExecuteNonQuery(reqSql);
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : UpdateLigneBASE() in DaoBob.cs.\n" + ex.Message);
                ok = false;
            }
            return ok;
        }

        public static Boolean UpdateLigneTVA(Bob50KlDbk Ligne, String dossier, String type = "ACH")
        {
            Boolean ok;
            string reqSql;
            try
            {
                ok = true;
                daoADS.Directory = ReadBobIni.LinkDirectory + @"\" + dossier;
                reqSql = $"UPDATE LINK_KL{type} SET " + 
                          "TVATAMN = " + Ligne.TVatAmn.ToString().Replace(",", ".") + ", " + 
                          "TVATTOTAMN = " + Ligne.TVatTotAmn.ToString().Replace(",", ".") + " " + 
                          "WHERE TDBK = '" + Ligne.Tdbk + "' AND TFYEAR = '" + Ligne.Tfyear + "' AND TMONTH = " + Ligne.Tmonth + " AND " + 
                                "TDOCNO = " + Ligne.Tdocno + " AND TDOCLINE = " + Ligne.Tdocline;
                daoADS.ExecuteNonQuery(reqSql);
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : UpdateLigneTVA() in DaoBob.cs.\n" + ex.Message);
                ok = false;
            }
            return ok;
        }

        public static Boolean AddLigne(Bob50KlDbk Ligne, List<String> listSecAna, String dossier, String type = "ACH")
        {
            String reqSql;
            Boolean ok;

            // AJOUT D'UNE LIGNE D'OD
            try
            {

                daoADS.Directory = ReadBobIni.LinkDirectory + @"\" + dossier;
                reqSql = $"INSERT INTO LINK_KL{type} (TDBK, TYEAR, TFYEAR, TDOCDATE, TMONTH, TDOCNO, TDOCLINE, TTYPELINE, TACTTYPE, TACCOUNT, TAMOUNT, TDC, " +
                                                            "TCURAMN, TREM, TCBVAT, TBASVAT, TVCTOTAMN, TVATTOTAMN, TCURVATAMN, TVATAMN, TVCDBLAMN, TVATDBLAMN, TBASLSTAMN, TVSTORED";
                for (int i = 0; i < listSecAna.Count; i++)
                {
                    reqSql = reqSql + ", COST_" + listSecAna[i];
                }
                reqSql = reqSql + ") VALUES ('" +
                                            Ligne.Tdbk + "', " +
                                            Ligne.Tyear + ", '" +
                                            Ligne.Tfyear + "', '" +
                                            Ligne.Tdocdate.ToString("yyyy-MM-dd") + "'," +
                                            Ligne.Tmonth + ", " +
                                            Ligne.Tdocno + ", " +
                                            Ligne.Tdocline + ", '" +
                                            Ligne.Ttypeline + "', '" +
                                            Ligne.Tacttype + "', '" +
                                            Ligne.Taccount + "', " +
                                            Ligne.Tamount.ToString().Replace(",", ".") + ", '" +
                                            Ligne.Tdc + "', " +
                                            Ligne.TCurAmn.ToString().Replace(",", ".") + ", '" +
                                            Ligne.TRem.Replace("'", "''") + "', " +
                                            Ligne.TCbVat.ToString().Replace(",", ".") + "," +
                                            Ligne.TBasVat.ToString().Replace(",", ".") + "," +
                                            Ligne.TVcTotAmn.ToString().Replace(",", ".") + "," +
                                            Ligne.TVatTotAmn.ToString().Replace(",", ".") + "," +
                                            Ligne.TCurVatAmn.ToString().Replace(",", ".") + "," +
                                            Ligne.TVatAmn.ToString().Replace(",", ".") + "," +
                                            Ligne.TVcDblAmn.ToString().Replace(",", ".") + "," +
                                            Ligne.TVatDblAmn.ToString().Replace(",", ".") + "," +
                                            Ligne.TBasLstAmn.ToString().Replace(",", ".") + ", '" +
                                            Ligne.TVStored + "'";


                for (int i = 0; i < listSecAna.Count; i++)
                {
                    switch (i)
                    {
                        case 0:
                            reqSql = reqSql + ", '" + Ligne.Cost_1 + "'";
                            break;
                        case 1:
                            reqSql = reqSql + ", '" + Ligne.Cost_2 + "'";
                            break;
                        case 2:
                            reqSql = reqSql + ", '" + Ligne.Cost_3 + "'";
                            break;
                        case 3:
                            reqSql = reqSql + ", '" + Ligne.Cost_4 + "'";
                            break;
                        case 4:
                            reqSql = reqSql + ", '" + Ligne.Cost_5 + "'";
                            break;
                        case 5:
                            reqSql = reqSql + ", '" + Ligne.Cost_6 + "'";
                            break;
                        case 6:
                            reqSql = reqSql + ",' " + Ligne.Cost_7 + "'";
                            break;
                        case 7:
                            reqSql = reqSql + ",' " + Ligne.Cost_8 + "'";
                            break;
                    }
                }
                reqSql = reqSql + ")";
                daoADS.ExecuteQuery(reqSql);
                ok = true;
            }
            catch (Exception ex)
            {
                ok = false;
                Log.WriteLog("DaoBob", "Erreur dans la fonction AddLigne() : " + ex.Message);
            }
            return ok;
        }

        public static TVAInfos GetVATInfos(string title, string type, string nat1, string compute, string dataPath)
        {
            DataTable table;
            TVAInfos infos = null;
            try
            {
                daoADS.Directory = dataPath;

                String query = "SELECT * FROM AC_VAT WHERE VTITLE = '" + title + "' AND VTYPE = '" + type + "' AND VNAT1 = '" + nat1 + "' AND VCOMPUTE = " + compute;
                table = daoADS.ExecuteQuery(query);
                if (table.Rows.Count == 1)
                {
                    infos = new TVAInfos();
                    infos.Vcompute = Convert.ToDouble(compute);
                    infos.Vdeduc = Convert.ToDouble(table.Rows[0]["VDEDUC"]);
                    infos.Vkind = table.Rows[0]["VKIND"].ToString();
                    infos.Vnat1 = nat1;
                    infos.Vsalpur = table.Rows[0]["VSALPUR"].ToString();
                    infos.Vstored = table.Rows[0]["VSTORED"].ToString();
                    infos.Vtype = type;
                    infos.Vtitle = title;
                    if (!DBNull.Value.Equals(table.Rows[0]["VOSS"]))
                        infos.Voss = Convert.ToBoolean(table.Rows[0]["VOSS"]);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : GetTVAInfos() in DaoBob.cs.\n" + ex.Message);
            }
            return infos;
        }

        public static Boolean AjoutCompan(Bob50Tiers tier, string dossier)
        {
            bool result = false;
            string reqSql;
            try
            {
                daoADS.Directory = dossier;
                reqSql = "INSERT INTO AC_COMPAN (CID, CCUSTYPE, CSUPTYPE, CNAME1, CADDRESS1, CZIPCODE, CLOCALITY, CVATCAT, CVATREF, CVATNO, CBANKORDERMANDATE, CBANKORDERPAY, CBANKORDERB2B, EMAILADDRESS) VALUES ('" +
                    tier.Cid.Replace("'", "''") + "','" + tier.Ccustype + "','" + tier.Csuptype + "','" + tier.Cname1.Replace("'", "''") + "','" + tier.Cadress1.Replace("'", "''") + "','" + tier.Czipcode + "','" + tier.Clocality.Replace("'", "''") + "','" +
                    tier.Cvatcat + "','" + tier.Cvatref + "','" + tier.Cvatno + "','" + tier.Cbankordermandate + "'," + tier.Cbankorderpay + "," + tier.Cbankorderb2b + ",'" + tier.Emailaddress + "')";
                Log.WriteLog("TraceDebug", "AjoutCompan - Qurey " + reqSql);
                daoADS.ExecuteQuery(reqSql);
                
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : AjoutCompan() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static Boolean UpdateCompan(Bob50Tiers tier, string type = "C", string dossier = "")
        {
            bool result = false;
            string reqSql = "";


            try
            {
                daoADS.Directory = dossier;
                reqSql = "UPDATE AC_COMPAN ";
                if (type == "C")
                    reqSql = reqSql + "SET CCUSTYPE = '" + tier.Ccustype + "'";
                else if(type == "S")
                    reqSql = reqSql + "SET CSUPTYPE = '" + tier.Csuptype + "'";

                reqSql += ", " +
                   "CNAME1 = '" + tier.Cname1.Replace("'", "''") + "', " +
                   "CADDRESS1 = '" + tier.Cadress1.Replace("'", "''") + "', " +
                   "CZIPCODE = '" + tier.Czipcode + "', " +
                   "CLOCALITY = '" + tier.Clocality.Replace("'", "''") + "', " +
                   "CVATCAT = '" + tier.Cvatcat + "', " +
                   "CVATREF = '" + tier.Cvatref + "', " +
                   "CVATNO = '" + tier.Cvatno + "'";

                // On ne modifie l'email que s'il y a une valeur
                if (!string.IsNullOrWhiteSpace(tier.Emailaddress))
                {
                    reqSql += ", EMAILADDRESS = '" +
                              tier.Emailaddress.Replace("'", "''") + "'";
                }

                reqSql += " WHERE CID = '" + tier.Cid.Replace("'", "''") + "'";

                daoADS.ExecuteQuery(reqSql);
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : UpdateCompan() in DaoBob.cs.\n " + reqSql + "\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static Boolean UpdateCompanBank(Bob50Tiers tier, string dossier)
        {
            bool result = false;
            string reqSql;
            try
            {
                daoADS.Directory = dossier;
                reqSql = "UPDATE AC_COMPAN SET CBANKORDERMANDATE = '" + tier.Cbankordermandate + "', " +
                    "CBANKORDERPAY = " + tier.Cbankorderpay + ", " +
                    "CBANKORDERB2B = " + tier.Cbankorderb2b + " WHERE CID = '" + tier.Cid.Replace("'", "''") + "'";

                daoADS.ExecuteQuery(reqSql);
                
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : UpdateCompanBank() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static Boolean AjoutCompBk(CompBk tier, string dossier)
        {
            bool result = false;
            string reqSql;
            try
            {
                daoADS.Directory = dossier;
                reqSql = "INSERT INTO AC_COMPBK(CID, BANKNOINDEX, ORDER, TYPE, CBANKNO, CBANKCODE, CFORBNKNO, CFORBNKSWIFT, CFORBNKCTRY, DEFAULT) VALUES('" +
                tier.Cid.Replace("'", "''") + "','" + tier.Banknoindex + "','" + tier.Order + "','" + tier.Type + "','" + tier.Cbankno + "','" + tier.Cbankiban + "','" + tier.Cbankcode + "','" +
                tier.Cforbnkno + "','" + tier.Cforbnkswift + "','" + tier.Cforbnkctry + "','" + tier.Cdefault + "')";

               daoADS.ExecuteQuery(reqSql);
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : AjoutUpdateCompBk() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }
            return result;
        }

        public static bool UpdateCompBk(CompBk compbk, string dossier)
        {
            bool result = false;
            string reqSql;
            try
            {
                daoADS.Directory = dossier;
                reqSql = "UPDATE AC_COMPBK SET " +
                    "TYPE = '" + compbk.Type + "', " +
                    "CFORBNKNO = '" + compbk.Cforbnkno + "', " +
                    "CFORBNKSWIFT = '" + compbk.Cforbnkswift + "', " +
                    "CFORBNKCTRY = '" + compbk.Cforbnkctry + "', " +
                    "DEFAULT = '" + compbk.Cdefault + "'" +
                    "WHERE CID = '" + compbk.Cid.Replace("'", "''") + "' AND BANKNOINDEX = '" + compbk.Banknoindex + "'" ;

                daoADS.ExecuteQuery(reqSql);
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : UpdateCompBk() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static bool UpdateCompBkDefault(String cid, String banknoindex, String def, string dossier)
        {
            bool result = false;
            string reqSql;
            try
            {
                daoADS.Directory = dossier;
                reqSql = "UPDATE AC_COMPBK SET " +
                    "DEFAULT = '" + def + "'" +
                    "WHERE CID = '" + cid.Replace("'", "''") + "' AND BANKNOINDEX != '" + banknoindex + "'";

                daoADS.ExecuteQuery(reqSql);
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : UpdateCompBkDefault() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static Boolean AjoutUpdateCompDe(CompDe tier, string dossier)
        {
            bool result = false;
            string reqSql;
            try
            {
                daoADS.Directory = dossier;
                reqSql = "MERGE INTO AC_COMPDE ON CID = " + tier.Cid.Replace("'", "''") +
                    "WHEN MATCHED THEN" +
                    "UPDATE SET CFORBNKNO = '" + tier.Cforbnkno + "', " +
                    "UPDATE SET CFORBNKSWIFT = '" + tier.Cforbnkswift + "', " +
                    "CFORBNKCTRY = '" + tier.Cforbnkctry +  "' " +
                    "WHEN NOT MATCHED THEN" +
                    "INSERT (CID, CFORBNKNO, CFORBNKSWIFT, CFORBNKCTRY) VALUES ('" +
                    tier.Cid.Replace("'", "''") + "','" + tier.Cforbnkno + "','" + tier.Cforbnkswift + "','" + tier.Cforbnkctry + "')";
                daoADS.ExecuteQuery(reqSql);
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : AjoutUpdateCompDe() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static Boolean UpdateTierBankInfos(string cid, string cbankno, string cbankiban, string cbankcode, string cbnktypepay, string dossier)
        {
            Boolean result = false;
            string reqSql;
            try {
                daoADS.Directory = dossier;
                reqSql = $"UPDATE AC_COMPAN SET CBANKNO='{cbankno}', CBANKIBAN='{cbankiban}', CBANKCODE='{cbankcode}', cbnktypepay='{cbnktypepay}' WHERE cid = '{cid.Replace("'", "''")}'";
                daoADS.ExecuteQuery(reqSql);
                result = true;
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : AjoutUpdateCompan() in DaoBob.cs.\n" + ex.Message);
                result = false;
            }

            return result;
        }

        public static void GetListPdfImported(String dossier, String type = "ACH")
        {
            String reqSql = "";
            try
            {
                daoADS.Directory = ReadBobIni.LinkDirectory + @"\" + dossier;
                reqSql = $"SELECT * FROM LINK_KH{type} WHERE ISNULL(TPDFFILENAME, '') <> ''"; 
                DataTable table = daoADS.ExecuteQuery(reqSql);
                foreach (DataRow row in table.Rows)
                {
                    VarGlobal.listPdf.Add(row["TPDFFILENAME"].ToString());
                }
            }
            catch (Exception ex)
            {
                Log.WriteLog("DaoBob", "Method : AddEntete() in DaoBob.cs.\n" + reqSql + "\n" + ex.Message);
            }
        }

        public static Bob50Config LOCAL_GetConfigDossier(Bob50Folder folder)
        {
            try
            {
                Bob50Config bob50Config = new Bob50Config();
                daoADS.Directory = folder.Path;
                DataTable dataTable = daoADS.ExecuteQuery("SELECT * FROM AC_CONFIG");
                bob50Config.Name = dataTable.Rows[0]["Name"].ToString();
                bob50Config.Name2 = dataTable.Rows[0]["Name2"].ToString();
                bob50Config.Addr1 = dataTable.Rows[0]["Addr1"].ToString();
                bob50Config.Addr2 = dataTable.Rows[0]["Addr2"].ToString();
                bob50Config.Zipcode = dataTable.Rows[0]["Zipcode"].ToString();
                bob50Config.Locality = dataTable.Rows[0]["Locality"].ToString();
                bob50Config.Phone = dataTable.Rows[0]["Phone"].ToString();
                bob50Config.Fax = dataTable.Rows[0]["Fax"].ToString();
                bob50Config.Email = dataTable.Rows[0]["Email"].ToString();
                bob50Config.Vatcode = dataTable.Rows[0]["Vatcode"].ToString();
                bob50Config.Vatno = dataTable.Rows[0]["Vatno"].ToString();
                bob50Config.Bankcode = dataTable.Rows[0]["Bankcode"].ToString();
                bob50Config.Bankno = dataTable.Rows[0]["Bankno"].ToString();
                bob50Config.Legis = dataTable.Rows[0]["Legis"].ToString();
                bob50Config.Language = dataTable.Rows[0]["Language"].ToString();
                bob50Config.Altlang = dataTable.Rows[0]["Altlang"].ToString();
                bob50Config.Basecurrid = dataTable.Rows[0]["Basecurrid"].ToString();
                bob50Config.Basecurrde = Convert.ToInt32(dataTable.Rows[0]["Basecurrde"]);
                bob50Config.Path = dataTable.Rows[0]["Path"].ToString();
                bob50Config.Periodtrim = Convert.ToBoolean(dataTable.Rows[0]["Periodtrim"]);
                bob50Config.Acctlength = Convert.ToInt32(dataTable.Rows[0]["Acctlength"]);
                bob50Config.Modcost = Convert.ToBoolean(dataTable.Rows[0]["Modcost"]);
                bob50Config.Modcost1 = Convert.ToBoolean(dataTable.Rows[0]["Modcost1"]);
                bob50Config.Modcost2 = Convert.ToBoolean(dataTable.Rows[0]["Modcost2"]);
                bob50Config.Modbank = Convert.ToBoolean(dataTable.Rows[0]["Modbank"]);
                bob50Config.Modbank1 = Convert.ToBoolean(dataTable.Rows[0]["Modbank1"]);
                bob50Config.Modbank2 = Convert.ToBoolean(dataTable.Rows[0]["Modbank2"]);
                bob50Config.Modbank3 = Convert.ToBoolean(dataTable.Rows[0]["Modbank3"]);
                bob50Config.Modinvedit = Convert.ToBoolean(dataTable.Rows[0]["Modinvedit"]);
                bob50Config.Prmvatenbld = Convert.ToBoolean(dataTable.Rows[0]["Prmvatenbld"]);
                bob50Config.Vatname = dataTable.Rows[0]["Vatname"].ToString();
                bob50Config.Vataddress = dataTable.Rows[0]["Vataddress"].ToString();
                bob50Config.Vatpersname = dataTable.Rows[0]["Vatpersname"].ToString();
                bob50Config.Vatperstitle = dataTable.Rows[0]["Vatperstitle"].ToString();
                bob50Config.Curmanlevel = Convert.ToInt32(dataTable.Rows[0]["Curmanlevel"]);
                bob50Config.Prmintraenbld = Convert.ToBoolean(dataTable.Rows[0]["Prmintraenbld"]);
                bob50Config.Ispers1 = dataTable.Rows[0]["Ispers1"].ToString();
                bob50Config.Isnocell = dataTable.Rows[0]["Isnocell"].ToString();
                bob50Config.Isludetailled = Convert.ToBoolean(dataTable.Rows[0]["Isludetailled"]);
                bob50Config.Isotherdeclar = Convert.ToBoolean(dataTable.Rows[0]["Isotherdeclar"]);
                bob50Config.Isname2 = dataTable.Rows[0]["Isname2"].ToString();
                bob50Config.Isaddress2 = dataTable.Rows[0]["Isaddress2"].ToString();
                bob50Config.Iszipcode2 = dataTable.Rows[0]["Iszipcode2"].ToString();
                bob50Config.Islocality2 = dataTable.Rows[0]["Islocality2"].ToString();
                bob50Config.Ispers2 = dataTable.Rows[0]["Ispers2"].ToString();
                bob50Config.Isphone2 = dataTable.Rows[0]["Isphone2"].ToString();
                bob50Config.Isfax2 = dataTable.Rows[0]["Isfax2"].ToString();
                bob50Config.Isvatctry2 = dataTable.Rows[0]["Isvatctry2"].ToString();
                bob50Config.Dispduedelay = dataTable.Rows[0]["Dispduedelay"].ToString();
                bob50Config.Prnduedelay = dataTable.Rows[0]["Prnduedelay"].ToString();
                bob50Config.Lyuptodate = Convert.ToBoolean(dataTable.Rows[0]["Lyuptodate"]);
                bob50Config.Expenseclasses = dataTable.Rows[0]["Expenseclasses"].ToString();
                bob50Config.Incomeclasses = dataTable.Rows[0]["Incomeclasses"].ToString();
                bob50Config.Costclasses = dataTable.Rows[0]["Costclasses"].ToString();
                bob50Config.Colleccus = Encoding.UTF8.GetString((byte[])dataTable.Rows[0]["Colleccus"]);
                bob50Config.Collecsup = Encoding.UTF8.GetString((byte[])dataTable.Rows[0]["Collecsup"]);
                bob50Config.Defacct = dataTable.Rows[0]["Defacct"].ToString();
                bob50Config.Rmndmdelay = (DBNull.Value.Equals(dataTable.Rows[0]["Rmndmdelay"]) ? 0.0 : Convert.ToDouble(dataTable.Rows[0]["Rmndmdelay"]));
                bob50Config.Rmndadelay = (DBNull.Value.Equals(dataTable.Rows[0]["Rmndadelay"]) ? 0.0 : Convert.ToDouble(dataTable.Rows[0]["Rmndadelay"]));
                bob50Config.Rmndmaxlev = (DBNull.Value.Equals(dataTable.Rows[0]["Rmndmaxlev"]) ? 0.0 : Convert.ToDouble(dataTable.Rows[0]["Rmndmaxlev"]));
                bob50Config.Vattrim = Convert.ToBoolean(dataTable.Rows[0]["Vattrim"]);
                bob50Config.Invstock = Convert.ToBoolean(dataTable.Rows[0]["Invstock"]);
                bob50Config.Invorders = Convert.ToBoolean(dataTable.Rows[0]["Invorders"]);
                bob50Config.Invupricedec = ((!DBNull.Value.Equals(dataTable.Rows[0]["Invupricedec"])) ? Convert.ToInt32(dataTable.Rows[0]["Invupricedec"]) : 0);
                bob50Config.Invqtydec = ((!DBNull.Value.Equals(dataTable.Rows[0]["Invqtydec"])) ? Convert.ToInt32(dataTable.Rows[0]["Invqtydec"]) : 0);
                bob50Config.Invtreeart = Convert.ToBoolean(dataTable.Rows[0]["Invtreeart"]);
                bob50Config.Invaltcurrid = dataTable.Rows[0]["Invaltcurrid"].ToString();
                bob50Config.Invcurrprlist = Convert.ToBoolean(dataTable.Rows[0]["Invcurrprlist"]);
                bob50Config.Invmultihead = Convert.ToBoolean(dataTable.Rows[0]["Invmultihead"]);
                bob50Config.Invmultisupp = Convert.ToBoolean(dataTable.Rows[0]["Invmultisupp"]);
                bob50Config.Invcomponent = Convert.ToBoolean(dataTable.Rows[0]["Invcomponent"]);
                bob50Config.Invdefvatn = dataTable.Rows[0]["Invdefvatn"].ToString();
                bob50Config.Invdefvati = dataTable.Rows[0]["Invdefvati"].ToString();
                bob50Config.Invdefvate = dataTable.Rows[0]["Invdefvate"].ToString();
                bob50Config.Invdefimputn = dataTable.Rows[0]["Invdefimputn"].ToString();
                bob50Config.Invdefimputi = dataTable.Rows[0]["Invdefimputi"].ToString();
                bob50Config.Invdefimpute = dataTable.Rows[0]["Invdefimpute"].ToString();
                bob50Config.Bnkpaydbk = dataTable.Rows[0]["Bnkpaydbk"].ToString();
                bob50Config.Bnkdomdbk = dataTable.Rows[0]["Bnkdomdbk"].ToString();
                bob50Config.Bnkindir = dataTable.Rows[0]["Bnkindir"].ToString();
                bob50Config.Bnkoutdir = dataTable.Rows[0]["Bnkoutdir"].ToString();
                bob50Config.Modfixed = Convert.ToBoolean(dataTable.Rows[0]["Modfixed"]);
                bob50Config.Specialtype = dataTable.Rows[0]["Specialtype"].ToString();
                bob50Config.Is19nat = dataTable.Rows[0]["Is19nat"].ToString();
                bob50Config.Is29nat = dataTable.Rows[0]["Is29nat"].ToString();
                bob50Config.Is19reg = dataTable.Rows[0]["Is19reg"].ToString();
                bob50Config.Is29reg = dataTable.Rows[0]["Is29reg"].ToString();
                bob50Config.Is19prov = dataTable.Rows[0]["Is19prov"].ToString();
                bob50Config.Is29dest = dataTable.Rows[0]["Is29dest"].ToString();
                bob50Config.Is19trans = dataTable.Rows[0]["Is19trans"].ToString();
                bob50Config.Is29trans = dataTable.Rows[0]["Is29trans"].ToString();
                bob50Config.Is19incoterm = dataTable.Rows[0]["Is19incoterm"].ToString();
                bob50Config.Is29incoterm = dataTable.Rows[0]["Is29incoterm"].ToString();
                bob50Config.Tradereg = dataTable.Rows[0]["Tradereg"].ToString();
                bob50Config.Businessno = dataTable.Rows[0]["Businessno"].ToString();
                bob50Config.Officevatiban = dataTable.Rows[0]["Officevatiban"].ToString();
                bob50Config.Officevatvcs = dataTable.Rows[0]["Officevatvcs"].ToString();
                bob50Config.Vatlettersiban = dataTable.Rows[0]["Vatlettersiban"].ToString();
                bob50Config.Vatletterspath = dataTable.Rows[0]["Vatletterspath"].ToString();
                bob50Config.Vatlettersround = Convert.ToBoolean(dataTable.Rows[0]["Vatlettersround"]);
                bob50Config.Refaccchart = dataTable.Rows[0]["Refaccchart"].ToString();
                bob50Config.Statementpath = dataTable.Rows[0]["Statementpath"].ToString();
                bob50Config.Statementarchive = dataTable.Rows[0]["Statementarchive"].ToString();
                bob50Config.Pfidid = dataTable.Rows[0]["Pfidid"].ToString();
                bob50Config.Vertsector = dataTable.Rows[0]["Vertsector"].ToString();
                bob50Config.Intracomtrim = Convert.ToBoolean(dataTable.Rows[0]["Intracomtrim"]);
                bob50Config.Dematsystem = dataTable.Rows[0]["Dematsystem"].ToString();
                bob50Config.Dmsarchive = Convert.ToBoolean(dataTable.Rows[0]["Dmsarchive"]);
                bob50Config.Invitationguid = dataTable.Rows[0]["Invitationguid"].ToString();
                bob50Config.Pfidactive = Convert.ToBoolean(dataTable.Rows[0]["Pfidactive"]);
                bob50Config.Acqnumtype = Convert.ToBoolean(dataTable.Rows[0]["Acqnumtype"]);
                bob50Config.Acqfile = Convert.ToBoolean(dataTable.Rows[0]["Acqfile"]);
                bob50Config.Acqlength = ((!DBNull.Value.Equals(dataTable.Rows[0]["Acqlength"])) ? Convert.ToInt32(dataTable.Rows[0]["Acqlength"]) : 0);
                bob50Config.Acqbatch = Convert.ToBoolean(dataTable.Rows[0]["Acqbatch"]);
                bob50Config.Edisage = dataTable.Rows[0]["Edisage"].ToString();
                bob50Config.Exportpath = dataTable.Rows[0]["Exportpath"].ToString();
                bob50Config.Dmssystem = dataTable.Rows[0]["Dmssystem"].ToString();
                bob50Config.Acqlabdbk = Convert.ToBoolean(dataTable.Rows[0]["Acqlabdbk"]);
                bob50Config.Sofiskpath = dataTable.Rows[0]["Sofiskpath"].ToString();
                bob50Config.Odrepdeffreq = dataTable.Rows[0]["Odrepdeffreq"].ToString();
                bob50Config.Odrepcalclmethod = dataTable.Rows[0]["Odrepcalclmethod"].ToString();
                bob50Config.Group = ((!DBNull.Value.Equals(dataTable.Rows[0]["Group"])) ? Convert.ToInt32(dataTable.Rows[0]["Group"]) : 0);
                bob50Config.Apikey = dataTable.Rows[0]["Apikey"].ToString();
                bob50Config.Ecom_integration = dataTable.Rows[0]["Ecom_integration"].ToString();
                bob50Config.Freedelitypromo = Convert.ToBoolean(dataTable.Rows[0]["Freedelitypromo"]);
                bob50Config.Vatpersemail = dataTable.Rows[0]["Vatpersemail"].ToString();
                bob50Config.Digilabels = ((!DBNull.Value.Equals(dataTable.Rows[0]["Digilabels"])) ? Convert.ToBoolean(dataTable.Rows[0]["Digilabels"]) : false);
                //bob50Config.Digilabels = Convert.ToBoolean(dataTable.Rows[0]["Digilabels"]);
                bob50Config.Contextname = dataTable.Rows[0]["Contextname"].ToString();
                bob50Config.Filedatecorrect = Convert.ToBoolean(dataTable.Rows[0]["Filedatecorrect"]);
                bob50Config.Costmand = Convert.ToBoolean(dataTable.Rows[0]["Costmand"]);
                return bob50Config;
            }
            catch (Exception ex)
            {
                ExceptionManager.WriteException(ex);
                throw ex;
            }
        }

    }
}
