using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Configuration;
using System.Windows.Forms;
using System.IO;
using D4L.Bob50ImportStandardExcelAchVen.Models;
using D4L.Bob50ImportStandardExcelAchVen.Controller;
using Dev4LogLayer.Globals.GUI;

namespace D4L.Bob50ImportStandardExcelAchVen.GUI
{
    public partial class ImportsExcel : Form
    {
        private ControllerImportExcel currController;

        public ImportsExcel(ControllerImportExcel controler)
        {
            currController = controler;
            InitializeComponent();
        }


        /// <summary>
        /// CLICK BOUTON QUITTER
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtQuitter_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        /// <summary>
        /// CLICK BOUTON SELECTIONNER FICHIER EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RechDossier_Click(object sender, EventArgs e)
        {
            currController.SelectFichierExcel();
        }


        /// <summary>
        /// CLICK BOUTON EXECUTER EN MODE STANDARD (DOSSIER COURANT)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtExecuter_Click(object sender, EventArgs e)
        {
            currController.RunImportStandard();
        }


        /// <summary>
        /// CLICK BOUTON EXECUTER EN MODE MULTI-DOSSIERS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtExecuterMulti_Click(object sender, EventArgs e)
        {
            currController.RunImportMultiDossiers();
        }


        /// <summary>
        /// MENU LOGS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuLogs_Click(object sender, EventArgs e)
        {
            currController.GestionDesLogs();
        }


        /// <summary>
        /// MENU LICENCE
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuLicence_Click(object sender, EventArgs e)
        {
            currController.GestionLicence();
        }

        /// <summary>
        /// LANCEMENT DE LA FENETRE DE CONFIGURATION MAIL DEV4LOG AZUR
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItemDev4LogMail_Click(object sender, EventArgs e)
        {
            currController.SettingsMailDev4LogAzur();
        }

    }
}