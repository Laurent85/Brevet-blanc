using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Brevet_blanc
{
    public partial class Principal : Form
    {
        public Principal()
        {
            InitializeComponent();
        }

        public System.Data.DataTable TableNotes = new System.Data.DataTable();
        public System.Data.DataTable TableComposantes = new System.Data.DataTable();
        public int RowCount;
        public string Classe;
        public string Progression;

        private void Principal_Load(object sender, EventArgs e)
        {
            TuerProcessus("Excel");
            lblSource.Text = @"X:\Logiciels\";
            lblDestination.Text = @"C:\Users\User\Desktop\";
            rdbSansOral.Checked = true;
            rdbDnb1.Checked = true;
            RemplirDatatable(TableNotes, lblSource.Text, "*.xls*", "Recapitulatif", "Notes", "AliasFichierNotes");
            RemplirDatatable(TableComposantes, lblSource.Text, "*.xls*", "Composantes", "Composantes", "AliasFichierComposantes");
            RemplirListeBox(chkLb_Notes, TableNotes);
            for (int i = 0; i < 5; i++)
                chkLb_Notes.SetItemChecked(i, true);
            RemplirListeBox(chkLb_Composantes, TableComposantes);
            for (int i = 0; i < 5; i++)
                chkLb_Composantes.SetItemChecked(i, true);
            chkLb_Notes_SelectedIndexChanged(sender, e);
        }

        private void BtnChoisirSource(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                lblSource.Text = dlg.SelectedPath + @"\";
            }

            Principal_Load(sender, e);
        }

        private void BtnChoisirDestination(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                lblDestination.Text = dlg.SelectedPath + @"\";
            }

            Directory.CreateDirectory(lblDestination.Text + @"DNB");
            Directory.CreateDirectory(lblDestination.Text + @"DNB\Composantes");
            Directory.CreateDirectory(lblDestination.Text + @"DNB\Modèles");
            Directory.CreateDirectory(lblDestination.Text + @"DNB\Notes");
        }

        private void BtnGénérerDnb(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            lblCompteur.Visible = true;
            ThreadDiplomes.RunWorkerAsync();
        }

        private void ThreadDiplomesMéthode(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            Directory.GetFiles(lblDestination.Text + @"DNB\", "*.*");

            var strPath = lblDestination.Text + @"DNB\Modèles\Dnb_sans_oral.xlsx";
            if (File.Exists(strPath)) File.Delete(strPath);
            var assembly = Assembly.GetExecutingAssembly();
            var input = assembly.GetManifestResourceStream("Brevet_blanc.Resources.Type_dnb.xlsx");
            var output = File.Open(strPath, FileMode.CreateNew);
            CopieFichiersTypeDnb(input, output);
            input?.Dispose();
            output.Dispose();

            var strPath1 = lblDestination.Text + @"DNB\Modèles\Dnb_sans_oral.docx";
            if (File.Exists(strPath1)) File.Delete(strPath1);
            var assembly1 = Assembly.GetExecutingAssembly();
            var input1 = assembly1.GetManifestResourceStream("Brevet_blanc.Resources.Type_dnb.docx");
            var output1 = File.Open(strPath1, FileMode.CreateNew);
            CopieFichiersTypeDnb(input1, output1);
            input1?.Dispose();
            output1.Dispose();

            #region TraductionComposantes

            string nomduFichierComposantes = "";
            int k;

            foreach (var fichierComposantes in chkLb_Composantes.CheckedItems)
            {
                foreach (DataRow ligne in TableComposantes.Rows)
                {
                    if (ligne[1].ToString() == fichierComposantes.ToString())
                    {
                        nomduFichierComposantes = ligne[0].ToString();  //Traduction du fichier date vers son vrai nom
                    }
                }
                var strPath2 = lblDestination.Text + @"DNB\Modèles\Dnb_sans_oral.xlsx";
                var appExcel1 = new Microsoft.Office.Interop.Excel.Application();
                var excelDocument1 = appExcel1.Workbooks.Open(strPath2);

                var récapitulatif = (Worksheet)excelDocument1.Sheets.Item[1];
                var épreuvesEcrites = (Worksheet)excelDocument1.Sheets.Item[2];

                var appExcel = new Microsoft.Office.Interop.Excel.Application();
                var excelDocument = appExcel.Workbooks.Open(lblSource.Text + nomduFichierComposantes);
                Worksheet worksheet = excelDocument.ActiveSheet;

                var a3 = (string)(worksheet.Cells[3, 1] as Range)?.Value;
                Classe = a3.Substring(0, 2);
                File.Delete(lblDestination.Text + @"DNB\" + NumDnb() + "-" + Classe + @".xlsx");
                File.Delete(lblDestination.Text + @"DNB\" + NumDnb() + "-" + Classe + @".pdf");
                File.Delete(lblDestination.Text + @"DNB\Statistiques.xlsx");
                File.Delete(lblDestination.Text + @"DNB\Statistiques.pdf");
                var effectifTemp = a3.Substring(a3.Length - 2);
                int effectif = int.Parse(effectifTemp);

                var début = worksheet.Cells[4, 1];
                var fin = worksheet.Cells[11, effectif + 2];
                var range = worksheet.Range[début, fin];

                RowCount = range.Cells.Count;
                k = 0;
                Progression = "Traduction couleurs composantes en points...";
                foreach (Range element in range.Cells) //Transformation des couleurs en points
                {
                    if (element.Font.ColorIndex == -5) //Rouge
                        element.Value2 = 10;
                    if (element.Font.ColorIndex == 1) //Orange
                        element.Value2 = 25;
                    if (element.Font.ColorIndex == 2) //Bleu
                        element.Value2 = 40;
                    if (element.Font.ColorIndex == 3) //Vert
                        element.Value2 = 50;
                    k++;
                    ThreadDiplomes.ReportProgress(k);
                }

                appExcel.DisplayAlerts = false;
                excelDocument.SaveAs(lblDestination.Text + @"DNB\DNB-" + Classe);

                #endregion TraductionComposantes

                #region CopieNomsEtComposantes

                var appExcel2 = new Microsoft.Office.Interop.Excel.Application();
                var excelDocument2 = appExcel2.Workbooks.Open(lblDestination.Text + @"DNB\DNB-" + Classe);
                Worksheet worksheet2 = excelDocument2.ActiveSheet;

                RowCount = effectif;
                k = 0;
                Progression = "Copie du nom des élèves et de la classe...";
                for (int i = 2; i <= effectif + 1; i++) //Copie du nom des élèves et de la classe
                {
                    récapitulatif.Cells[i, 1].Value = worksheet.Cells[3, i].Value.ToString();
                    récapitulatif.Cells[i, 2].Value = Classe;
                    épreuvesEcrites.Cells[i, 1].Value = worksheet.Cells[3, i].Value.ToString();
                    k++;
                    ThreadDiplomes.ReportProgress(k);
                }

                RowCount = effectif;
                k = 0;
                Progression = "Copie des points des composantes...";
                for (int i = 2; i <= effectif + 1; i++) //Copie des points des composantes
                {
                    for (int j = 3; j <= 10; j++)
                    {
                        if (worksheet2.Cells[j + 1, i].Value != null)
                            récapitulatif.Cells[i, j].Value = worksheet2.Cells[j + 1, i].Value.ToString();
                    }
                    k++;
                    ThreadDiplomes.ReportProgress(k);
                }

                var cells = récapitulatif.Range["A" + (effectif + 2) + ":A500"]; //Nettoyage bas tableau récapitulatif
                var del = cells.EntireRow;
                del.Delete();

                appExcel1.DisplayAlerts = false;
                excelDocument1.SaveAs(lblDestination.Text + @"DNB\" + NumDnb() + "-" + Classe);

                excelDocument2.Close(0);
                excelDocument.Close(0);
                excelDocument1.Close(0);

                appExcel1.Quit();
                appExcel.Quit();
                appExcel2.Quit();

                Marshal.ReleaseComObject(appExcel);
                Marshal.ReleaseComObject(appExcel1);
                Marshal.ReleaseComObject(appExcel2);
            }

            #endregion CopieNomsEtComposantes

            #region CopieNotes

            string nomduFichierNotes = "";

            foreach (var fichierNotes in chkLb_Notes.CheckedItems)
            {
                foreach (DataRow ligne in TableNotes.Rows)
                {
                    if (ligne[1].ToString() == fichierNotes.ToString())
                    {
                        nomduFichierNotes = ligne[0].ToString();
                    }
                }
                var appExcel = new Microsoft.Office.Interop.Excel.Application();
                var excelDocument = appExcel.Workbooks.Open(lblSource.Text + nomduFichierNotes);
                Worksheet worksheet = excelDocument.ActiveSheet;

                var b1 = (string)(worksheet.Cells[1, 2] as Range)?.Value;
                Classe = b1.Substring(0, 2);
                var effectifTemp = b1.Substring(b1.Length - 2);
                int effectif = int.Parse(effectifTemp);

                appExcel.DisplayAlerts = false;
                var appExcel2 = new Microsoft.Office.Interop.Excel.Application();
                var excelDocument2 = appExcel2.Workbooks.Open(lblDestination.Text + @"DNB\" + NumDnb() + "-" + Classe);
                var récapitulatif = (Worksheet)excelDocument2.Sheets.Item[1];
                var épreuvesEcrites2 = (Worksheet)excelDocument2.Sheets.Item[2];

                RowCount = effectif;
                k = 0;
                Progression = "Copie des notes...";
                for (int i = 2; i <= effectif + 1; i++) //Copie des notes
                {
                    int m = 11;
                    int totalPoints = 0;
                    for (int j = 2; j <= 8; j++)
                    {
                        if (worksheet.Cells[i + 3, j + 9].Value != null)
                        {
                            épreuvesEcrites2.Cells[i, j].Value = worksheet.Cells[i + 3, j + 9].Value;
                            m = m + 1;
                            if (m == 14) m = 15;
                            totalPoints = totalPoints + Bareme(j);
                        }
                        if (worksheet.Cells[i + 3, j + 9].Value == null) //Gestion des cellules vides
                        {
                            récapitulatif.Cells[i, j + m].value = "";
                            épreuvesEcrites2.Cells[i, 1].value = épreuvesEcrites2.Cells[i, 1].value + "*";
                            m = m + 1;
                            if (m == 14) m = 15;
                        }
                    }
                    récapitulatif.Range["AR" + i].Value = totalPoints + 400;
                    récapitulatif.Range["AT" + i].Value = totalPoints;

                    k++;
                    ThreadDiplomes.ReportProgress(k);
                }

                var cells = épreuvesEcrites2.Range["A" + (effectif + 2) + ":A500"]; //Nettoyage bas tableau récapitulatif
                var del = cells.EntireRow;
                del.Delete();

                appExcel2.DisplayAlerts = false;
                excelDocument2.SaveAs(lblDestination.Text + @"DNB\" + NumDnb() + "-" + Classe);
                excelDocument2.Close(0);
                excelDocument.Close(0);
                appExcel.Quit();
                appExcel2.Quit();

                Marshal.ReleaseComObject(appExcel);
                Marshal.ReleaseComObject(appExcel2);
            }

            #endregion CopieNotes

            TuerProcessus("Excel");

            #region PublipostageDNB

            var fichiersDnb = Directory.GetFiles(lblDestination.Text + @"DNB\", "*.*");
            RowCount = 0;
            k = 0;
            Progression = "Publipostage des diplômes...";
            foreach (var unused in chkLb_Notes.CheckedItems)
            {
                RowCount++;
            }

            foreach (var fichier1 in chkLb_Notes.CheckedItems)
            {
                var classe = fichier1.ToString().Substring(0, 2);
                foreach (var fichier in fichiersDnb)
                {
                    if ((fichier.Contains(fichier1.ToString().Substring(0, 2))) && (fichier.Contains(NumDnb())))
                    {
                        var appWord = new Microsoft.Office.Interop.Word.Application();
                        var wordDocument = appWord.Documents.Add(lblDestination.Text + @"DNB\Modèles\Dnb_sans_oral.docx");
                        appWord.Visible = false;
                        wordDocument.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
                        var nomDuFichierDnb = Path.GetFileNameWithoutExtension(fichier);
                        Classe = nomDuFichierDnb.Substring(5, 2);

                        string strDataFile = fichier;
                        object objTrue = true;
                        object objFalse = false;
                        object objMiss = Missing.Value;
                        object type = WdMergeSubType.wdMergeSubTypeAccess;
                        object strQuery = "SELECT * FROM [Récapitulatif$]";
                        object connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strDataFile + ";Extended Properties=\"HDR=YES;IMEX=1\";Jet OLEDB:EngineType=37";

                        wordDocument.MailMerge.OpenDataSource(fichier, objMiss, objFalse, objTrue, objTrue, objFalse,
                            objMiss, objMiss, objMiss, objMiss, objMiss, connectionString, strQuery, objMiss, objMiss, type);

                        wordDocument.MailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                        wordDocument.MailMerge.SuppressBlankLines = true;
                        wordDocument.MailMerge.Execute(false);

                        var oLetters = appWord.ActiveDocument;
                        //oLetters.SaveAs2(@"F:\DNB\" +  nomDuFichierDnb + ".docx",
                        //WdSaveFormat.wdFormatDocumentDefault);
                        oLetters.ExportAsFixedFormat(lblDestination.Text + @"DNB\" + nomDuFichierDnb + ".pdf",
                            WdExportFormat.wdExportFormatPDF);
                        oLetters.Close(WdSaveOptions.wdDoNotSaveChanges);
                        wordDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
                        appWord.Quit();
                        GC.Collect();

                        k++;
                        ThreadDiplomes.ReportProgress(k);
                    }
                }

                if (File.Exists(lblDestination.Text + @"DNB\Composantes\DNB-" + classe + @".xlsx"))
                    File.Delete(lblDestination.Text + @"DNB\Composantes\DNB-" + classe + @".xlsx");
                File.Move(lblDestination.Text + @"DNB\DNB-" + classe + @".xlsx", lblDestination.Text + @"DNB\Composantes\DNB-" + classe + @".xlsx");

                if (File.Exists(lblDestination.Text + @"DNB\Notes\" + NumDnb() + "-" + classe + @".xlsx"))
                    File.Delete(lblDestination.Text + @"DNB\Notes\" + NumDnb() + "-" + classe + @".xlsx");
                File.Move(lblDestination.Text + @"DNB\" + NumDnb() + "-" + classe + @".xlsx", lblDestination.Text + @"DNB\Notes\" + NumDnb() + "-" + classe + @".xlsx");
            }

            TuerProcessus("Winword");

            #endregion PublipostageDNB
        }

        private void ThreadDiplomesProgression(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Maximum = RowCount;
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBar1.Value = e.ProgressPercentage;
            // Set the text.
            lblCompteur.Text = Progression + Environment.NewLine + Environment.NewLine + @"            " + e.ProgressPercentage + @" / " + RowCount;
            lblClasse.Text = @"Traitement des " + Classe;
        }

        private void ThreadDiplomesTerminé(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 0;
            lblCompteur.Text = "";
            lblClasse.ForeColor = System.Drawing.Color.ForestGreen;
            lblClasse.Text = @"Terminé !";
            lblClasses.ForeColor = System.Drawing.Color.ForestGreen;
        }

        private void BtnGénérerStatistiques(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            lblCompteur.Visible = true;
            ThreadStatistiques.RunWorkerAsync();
        }

        private void ThreadStatistiquesMéthode(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            #region Initialisation des classeurs Excel
            var fichiersDnbXlsx = Directory.GetFiles(lblDestination.Text + @"DNB\Notes\");
            var fichierStat = lblDestination.Text + @"DNB\Statistiques.xlsx";
            int k;
            if (!File.Exists(fichierStat))
            {
                var assembly = Assembly.GetExecutingAssembly();
                var input = assembly.GetManifestResourceStream("Brevet_blanc.Resources.Statistiques.xlsx");
                var output = File.Open(fichierStat, FileMode.CreateNew);
                CopieFichiersTypeDnb(input, output);
                input?.Dispose();
                output.Dispose();
            }

            var excelApplication = new Microsoft.Office.Interop.Excel.Application();
            var statXlsx = excelApplication.Workbooks.Open(fichierStat);
            var statSynthèse = (Worksheet)statXlsx.Sheets.Item[1];
            var statMoyennes = (Worksheet)statXlsx.Sheets.Item[2];
            var statMoyennesControle = (Worksheet)statXlsx.Sheets.Item[3];
            var statListing = (Worksheet)statXlsx.Sheets.Item[4];
            var statDelta = (Worksheet)statXlsx.Sheets.Item[5];
            #endregion Initialisation des classeurs Excel

            foreach (RadioButton dnb in panel1.Controls) //DNB1 et DNB2
            {
                #region Initialisation des variables (lignes)
                int nombreClasses = 0;
                int ligneStatSynthèse = 3;
                int ligneStatSynthèseDébut = 0;
                int ligneStatMoyennesEe = 3;
                int ligneStatMoyennesEeDébut = 0;
                int ligneStatMoyennesCc = 3;
                int ligneStatMoyennesCcDébut = 0;
                int ligneStatListingColA = 3;
                int ligneStatListingColE = 3;
                int ligneStatDelta = 3;
                int ligneEleve = 2;

                if (dnb.Text == "DNB1")
                {
                    ligneStatSynthèse = 3;
                    ligneStatMoyennesEe = 3;
                    ligneStatMoyennesCc = 3;
                    statListing = (Worksheet)statXlsx.Sheets.Item[4];
                    statDelta = (Worksheet)statXlsx.Sheets.Item[5];
                }
                if (dnb.Text == "DNB2")
                {
                    ligneStatSynthèse = 15;
                    ligneStatMoyennesEe = 15;
                    ligneStatMoyennesCc = 15;
                    statListing = (Worksheet)statXlsx.Sheets.Item[6];
                    statDelta = (Worksheet)statXlsx.Sheets.Item[7];
                }
                #endregion Initialisation des variables (lignes)

                #region Dnb1SynthèseEtListing

                #region Effacement Listing et Delta
                for (int i = 3; i < 33; i++) //Effacement Listing et Delta
                {
                    statDelta.Range["A" + i].Value = "";
                    statDelta.Range["E" + i].Value = "";
                    statListing.Range["A" + i].Value = "";
                    statListing.Range["E" + i].Value = "";
                }
                #endregion Effacement Listing et Delta

                RowCount = fichiersDnbXlsx.Count();
                k = 0;
                Progression = "Synthèse...";
                foreach (var file in fichiersDnbXlsx)
                {
                    #region initialisation des variables
                    var fichierDnbXlsx = Path.GetFileName(file);
                    if (fichierDnbXlsx.Contains(dnb.Text) && fichierDnbXlsx.Contains("xlsx")) //DNB1 ou DNB2
                    {
                        if ((ligneStatSynthèse == 3) || (ligneStatSynthèse == 15))
                        {
                            statSynthèse.Range["B" + (ligneStatSynthèse + 1) + ":G" + (ligneStatSynthèse + 10)].Value = 0;
                            ligneStatSynthèseDébut = ligneStatSynthèse;
                        }
                        ligneStatSynthèse++;
                        var fichierDnb = lblDestination.Text + @"DNB\Notes\" + fichierDnbXlsx;
                        var dnbXlsx = excelApplication.Workbooks.Open(fichierDnb);
                        var dnbRécapitulatif = (Worksheet)dnbXlsx.Sheets.Item[1];
                        var épreuvesEcrites = (Worksheet)dnbXlsx.Sheets.Item[2];

                        var range = dnbRécapitulatif.Range["AG2:AG50"];

                        statSynthèse.Range["A" + ligneStatSynthèse].Value = dnbRécapitulatif.Range["B2"].Value.ToString();
                        #endregion initialisation des variables

                        foreach (Range element in range.Cells)
                        {
                            # region Gestion des mentions pour Synthèse, Listing et Delta
                            if (element.Value2 != null)
                            {
                                statSynthèse.Range["B" + ligneStatSynthèse].Value =
                                        int.Parse(statSynthèse.Range["B" + ligneStatSynthèse].Value.ToString()) + 1;

                                if (element.Value.ToString().Contains("Non"))
                                {
                                    statSynthèse.Range["C" + ligneStatSynthèse].Value =
                                        int.Parse(statSynthèse.Range["C" + ligneStatSynthèse].Value.ToString()) + 1;
                                    statListing.Range["A" + ligneStatListingColA].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + épreuvesEcrites.Range["A" + ligneEleve].Value.ToString() + " (" + dnbRécapitulatif.Range["AG" + ligneEleve].Value.ToString() + ")";
                                    ligneStatListingColA++;
                                    int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value / 2 -
                                                dnbRécapitulatif.Range["AE" + ligneEleve].Value);

                                    if ((delta <= numDelta.Value) && (ligneStatDelta > 31))
                                    {
                                        statDelta.Range["E" + (ligneStatDelta - 29)].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour obtention)";
                                        ligneStatDelta++;
                                    }
                                    if ((delta <= numDelta.Value) && (ligneStatDelta <= 31))
                                    {
                                        statDelta.Range["A" + ligneStatDelta].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour obtention)";
                                        ligneStatDelta++;
                                    }
                                }
                                if (element.Value.ToString().Contains("sans mention"))
                                {
                                    statSynthèse.Range["D" + ligneStatSynthèse].Value =
                                        int.Parse(statSynthèse.Range["D" + ligneStatSynthèse].Value.ToString()) + 1;

                                    int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value * 12 / 20 -
                                                dnbRécapitulatif.Range["AE" + ligneEleve].Value);

                                    if ((delta <= numDelta.Value) && (ligneStatDelta > 31))
                                    {
                                        statDelta.Range["E" + (ligneStatDelta - 29)].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention AB)";
                                        ligneStatDelta++;
                                    }
                                    if ((delta <= numDelta.Value) && (ligneStatDelta <= 31))
                                    {
                                        statDelta.Range["A" + ligneStatDelta].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention AB)";
                                        ligneStatDelta++;
                                    }
                                }
                                if (element.Value.ToString().Contains("mention AB"))
                                {
                                    statSynthèse.Range["E" + ligneStatSynthèse].Value =
                                        int.Parse(statSynthèse.Range["E" + ligneStatSynthèse].Value.ToString()) + 1;

                                    int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value * 14 / 20 -
                                                dnbRécapitulatif.Range["AE" + ligneEleve].Value);

                                    if ((delta <= numDelta.Value) && (ligneStatDelta > 31))
                                    {
                                        statDelta.Range["E" + (ligneStatDelta - 29)].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention B)";
                                        ligneStatDelta++;
                                    }
                                    if ((delta <= numDelta.Value) && (ligneStatDelta <= 31))
                                    {
                                        statDelta.Range["A" + ligneStatDelta].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention B)";
                                        ligneStatDelta++;
                                    }
                                }
                                if (element.Value.ToString().Contains("mention B"))
                                {
                                    statSynthèse.Range["F" + ligneStatSynthèse].Value =
                                        int.Parse(statSynthèse.Range["F" + ligneStatSynthèse].Value.ToString()) + 1;

                                    int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value * 16 / 20 -
                                                dnbRécapitulatif.Range["AE" + ligneEleve].Value);

                                    if ((delta <= numDelta.Value) && (ligneStatDelta > 31))
                                    {
                                        statDelta.Range["E" + (ligneStatDelta - 29)].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention TB)";
                                        ligneStatDelta++;
                                    }
                                    if ((delta <= numDelta.Value) && (ligneStatDelta <= 31))
                                    {
                                        statDelta.Range["A" + ligneStatDelta].Value =
                                            dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention TB)";
                                        ligneStatDelta++;
                                    }
                                }
                                if (element.Value.ToString().Contains("mention TB"))
                                {
                                    statSynthèse.Range["G" + ligneStatSynthèse].Value =
                                        int.Parse(statSynthèse.Range["G" + ligneStatSynthèse].Value.ToString()) + 1;
                                    statListing.Range["E" + ligneStatListingColE].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + " (" + dnbRécapitulatif.Range["AE" + ligneEleve].Value.ToString() + " / " + dnbRécapitulatif.Range["AR" + ligneEleve].Value.ToString() + ")";
                                    ligneStatListingColE++;
                                }

                                statSynthèse.Range["H" + ligneStatSynthèse].Formula = "=SUM(D" + ligneStatSynthèse + ":G" + ligneStatSynthèse + ")";
                                statSynthèse.Range["I" + ligneStatSynthèse].Formula = "=H" + ligneStatSynthèse + "/B" + ligneStatSynthèse;
                            }
                            if (ligneEleve == 50) ligneEleve = 2;
                            else
                                ligneEleve++;

                            #endregion Dnb1SynthèseEtListing
                        }

                        #region Calcul colonne J - Moyennes générales
                        float total = 0;
                        int compteur = 0;
                        var colMoyennes = dnbRécapitulatif.Range["AF2:AF50"];
                        foreach (Range element in colMoyennes.Cells)
                        {
                            if (element.Value2 != null)
                            {
                                total = total + float.Parse(element.Value.ToString());
                                compteur++;
                            }
                        }
                        statSynthèse.Range["J" + ligneStatSynthèse].Value = total / compteur;
                        dnbXlsx.Close();
                        nombreClasses++;
                        #endregion Calcul colonne J - Moyennes générales
                    }
                    k++;
                    ThreadStatistiques.ReportProgress(k);
                }

                #region Nettoyage des cellules
                var range1 = statSynthèse.Range["A1:A1"];
                if (ligneStatSynthèse < 15)
                {
                    range1 = statSynthèse.Range["A" + (ligneStatSynthèse + 1),
                         "J13"];
                }
                if (ligneStatSynthèse > 15)
                {
                    range1 = statSynthèse.Range["A" + (ligneStatSynthèse + 1),
                         "J25"];
                }
                if ((ligneStatSynthèse == 3) || (ligneStatSynthèse == 15)) //Effacement des cellules
                {
                    statSynthèse.Range["A" + (ligneStatSynthèse + 1), "J" + (ligneStatSynthèse + 10)].Value = "";
                }
                else range1.Value = "";
                #endregion Nettoyage des cellules

                #region Calcul ligne "Niveau"
                if (nombreClasses > 0)
                {
                    statSynthèse.Range["A" + (ligneStatSynthèse + 2)].Value = "Niveau";
                    statDelta.Range["A2"].Value = ligneStatDelta - 3 + " élèves à " + numDelta.Value + " points ou moins pour atteindre un palier";

                    var colonne = 'B';
                    for (int i = 1; i < 8; i++) // Somme des candidats de la ligne "niveau"
                    {
                        statSynthèse.Range[colonne.ToString() + (ligneStatSynthèse + 2)].Formula =
                            "=SUM(" + colonne + ligneStatSynthèseDébut + ":" + colonne + ligneStatSynthèse + ")";
                        colonne++;
                    }

                    // Pourcentage et moyenne pour la ligne "niveau"
                    statSynthèse.Range["I" + (ligneStatSynthèse + 2)].Formula =
                        "=H" + (ligneStatSynthèse + 2) + "/B" + (ligneStatSynthèse + 2);
                    statSynthèse.Range["J" + (ligneStatSynthèse + 2)].Formula = "=AVERAGE(J" + ligneStatSynthèseDébut + ":J" + ligneStatSynthèse;
                }
                #endregion Calcul ligne "Niveau"

                #endregion Dnb1SynthèseEtListing

                #region Dnb1MoyennesEpreuves

                RowCount = fichiersDnbXlsx.Count();
                k = 0;
                Progression = "Epreuves écrites...";
                foreach (var file in fichiersDnbXlsx)
                {
                    var fichierDnbXlsx = Path.GetFileName(file);

                    if (fichierDnbXlsx.Contains(dnb.Text) && fichierDnbXlsx.Contains("xlsx"))
                    {
                        #region Initialisation des variables
                        if ((ligneStatMoyennesEe == 3) || (ligneStatMoyennesEe == 15))
                        {
                            statMoyennes.Range["B" + (ligneStatMoyennesEe + 1) + ":G" + (ligneStatMoyennesEe + 10)].Value = 0;
                            ligneStatMoyennesEeDébut = ligneStatMoyennesEe;
                        }
                        ligneStatMoyennesEe++;
                        var fichierDnb = lblDestination.Text + @"DNB\Notes\" + fichierDnbXlsx;
                        var dnbXlsx = excelApplication.Workbooks.Open(fichierDnb);
                        var dnbEpreuvesEcrites = (Worksheet)dnbXlsx.Sheets.Item[2];
                        statMoyennes.Range["A" + ligneStatMoyennesEe].Value = ((Worksheet)dnbXlsx.Sheets.Item[1]).Range["B2"].Value.ToString(); //classe
                        statMoyennes.Range["I" + ligneStatMoyennesEe].Value = ""; //oral
                        int effectif = int.Parse(statSynthèse.Range["B" + ligneStatMoyennesEe.ToString()].Value.ToString()); //effectif
                        #endregion
                        #region Calcul des moyennes par épreuve
                        var colonne = 'B';
                        for (int i = 1; i < 8; i++)
                        {
                            int barême = int.Parse(dnbEpreuvesEcrites.Range[colonne + "1"].Value.ToString().Split(new[] { '/', ')' })[1]);

                            dnbEpreuvesEcrites.Range[colonne.ToString() + (effectif + 3)].Formula = "=AVERAGE(" + colonne.ToString() + "2:" + colonne.ToString() + (effectif + 2) + ")";

                            statMoyennes.Range[colonne.ToString() + ligneStatMoyennesEe].Value = Math.Round(
                                float.Parse(dnbEpreuvesEcrites.Range[colonne.ToString() + (effectif + 3)].Value.ToString()) / barême * 20, 2);

                            colonne++;
                        }
                        #endregion
                        #region Calcul de la moyenne générale des épreuves
                        dnbEpreuvesEcrites.Range["J" + (effectif + 3)].Formula = "=AVERAGE(B" + (effectif + 3) + ":H" + (effectif + 3) + ")";
                        statMoyennes.Range["J" + ligneStatMoyennesEe].Value = Math.Round(float.Parse(dnbEpreuvesEcrites.Range["J" + (effectif + 3)].Value.ToString()), 2);
                        #endregion
                        object misValue = Missing.Value;
                        dnbXlsx.Close(false, misValue, misValue);
                    }
                    k++;
                    ThreadStatistiques.ReportProgress(k);
                }
                #region Effacement des cellules inutiles
                var range3 = statMoyennes.Range["A1:A1"];
                if (ligneStatMoyennesEe < 15)
                {
                    range3 = statMoyennes.Range["A" + (ligneStatMoyennesEe + 1),
                         "J13"];
                }
                if (ligneStatMoyennesEe > 15)
                {
                    range3 = statMoyennes.Range["A" + (ligneStatMoyennesEe + 1),
                         "J25"];
                }
                if ((ligneStatMoyennesEe == 3) || (ligneStatMoyennesEe == 15))
                {
                    statMoyennes.Range["A" + (ligneStatMoyennesEe + 1), "J" + (ligneStatMoyennesEe + 10)].Value = "";
                }
                else range3.Value = "";
                #endregion

                if (nombreClasses > 0)
                {
                    statMoyennes.Range["A1"].Value = "Année scolaire 2018-2019";
                    #region Calcul de la moyenne générale par épreuve pour le niveau
                    statMoyennes.Range["A" + (ligneStatMoyennesEe + 2)].Value = "Niveau";
                    var colonne = 'B';
                    for (int i = 1; i < 8; i++)
                    {
                        statMoyennes.Range[colonne.ToString() + (ligneStatMoyennesEe + 2)].Formula =
                            "=AVERAGE(" + colonne + ligneStatMoyennesEeDébut + ":" + colonne + ligneStatMoyennesEe + ")";
                        colonne++;
                    }
                    #endregion
                    #region Calcul de la moyenne générale pour le niveau
                    statMoyennes.Range["J" + (ligneStatMoyennesEe + 2)].Formula =
                        "=AVERAGE(J" + ligneStatMoyennesEeDébut + ":J" + ligneStatMoyennesEe + ")";
                    #endregion
                }

                #endregion Dnb1MoyennesEpreuves

                #region Dnb1MoyennesControleContinu

                RowCount = fichiersDnbXlsx.Count();
                k = 0;
                Progression = "Contrôle continu...";
                foreach (var file in fichiersDnbXlsx)
                {
                    var fichierDnbXlsx = Path.GetFileName(file);

                    if (fichierDnbXlsx.Contains(dnb.Text) && fichierDnbXlsx.Contains("xlsx"))
                    {
                        #region Initialisation des variables
                        if ((ligneStatMoyennesCc == 3) || (ligneStatMoyennesCc == 15))
                        {
                            statMoyennesControle.Range["B" + (ligneStatMoyennesCc + 1) + ":G" + (ligneStatMoyennesCc + 10)].Value = 0;
                            ligneStatMoyennesCcDébut = ligneStatMoyennesCc;
                        }
                        ligneStatMoyennesCc++;
                        var fichierDnb = lblDestination.Text + @"DNB\Notes\" + fichierDnbXlsx;
                        var dnbXlsx = excelApplication.Workbooks.Open(fichierDnb);
                        var dnbRécapitulatif = (Worksheet)dnbXlsx.Sheets.Item[1];
                        statMoyennesControle.Range["A" + ligneStatMoyennesCc].Value = ((Worksheet)dnbXlsx.Sheets.Item[1]).Range["B2"].Value.ToString();
                        statMoyennesControle.Range["I" + ligneStatMoyennesCc].Value = "";
                        int effectif = int.Parse(statSynthèse.Range["B" + ligneStatMoyennesCc.ToString()].Value.ToString());
                        #endregion
                        #region Calcul des moyennes par domaine
                        var colonne = 'B';
                        var colonne1 = 'C';
                        for (int i = 1; i < 9; i++)
                        {
                            float somme = 0;
                            for (int j = 2; j <= effectif + 1; j++)
                            {
                                if (dnbRécapitulatif.Range[colonne1.ToString() + j].Value != null)
                                    somme = somme + float.Parse(dnbRécapitulatif.Range[colonne1.ToString() + j].Value.ToString());
                            }

                            statMoyennesControle.Range[colonne.ToString() + ligneStatMoyennesCc].Value = somme / effectif;

                            colonne++;
                            colonne1++;
                        }
                        #endregion
                        #region Calcul de la moyenne générale des domaines
                        statMoyennesControle.Range["J" + ligneStatMoyennesCc].Formula = "=AVERAGE(B" + ligneStatMoyennesCc + ":I" + ligneStatMoyennesCc;
                        dnbXlsx.Close();
                        #endregion
                    }
                    k++;
                    ThreadStatistiques.ReportProgress(k);
                }
                #region Effacement des cellules inutiles
                var range5 = statMoyennesControle.Range["A1:A1"];
                if (ligneStatMoyennesCc < 15)
                {
                    range5 = statMoyennesControle.Range["A" + (ligneStatMoyennesCc + 1),
                         "J13"];
                }
                if (ligneStatMoyennesCc > 15)
                {
                    range5 = statMoyennesControle.Range["A" + (ligneStatMoyennesCc + 1),
                         "J25"];
                }

                if ((ligneStatMoyennesCc == 3) || (ligneStatMoyennesCc == 15))
                {
                    statMoyennesControle.Range["A" + (ligneStatMoyennesCc + 1), "J" + (ligneStatMoyennesCc + 10)].Value = "";
                }
                else range5.Value = "";
                #endregion
                if (nombreClasses > 0)
                {
                    statMoyennesControle.Range["A1"].Value = "Année scolaire 2018-2019";
                    #region Calcul de la moyenne générale par épreuve pour le niveau
                    statMoyennesControle.Range["A" + (ligneStatMoyennesCc + 2)].Value = "Niveau";

                    var colonne = 'B';
                    for (int i = 1; i < 9; i++)
                    {
                        statMoyennesControle.Range[colonne.ToString() + (ligneStatMoyennesCc + 2)].Formula =
                            "=AVERAGE(" + colonne + ligneStatMoyennesCcDébut + ":" + colonne + ligneStatMoyennesCc + ")";
                        colonne++;
                    }
                    #endregion
                    #region Calcul de la moyenne générale pour le niveau
                    statMoyennesControle.Range["J" + (ligneStatMoyennesCc + 2)].Formula =
                        "=AVERAGE(J" + ligneStatMoyennesCcDébut + ":J" + ligneStatMoyennesCc + ")";
                    #endregion
                }
            }

            #endregion Dnb1MoyennesControleContinu

            excelApplication.DisplayAlerts = false;
            statXlsx.SaveAs(fichierStat);
            statXlsx.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                            lblDestination.Text + @"DNB\Statistiques.pdf");
            statXlsx.Close();
            GC.Collect();
        }

        private void ThreadStatistiquesProgression(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Maximum = RowCount;
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBar1.Value = e.ProgressPercentage;
            // Set the text.
            lblCompteur.Text = Progression + Environment.NewLine + Environment.NewLine + @"            " + e.ProgressPercentage + @" / " + RowCount;
            lblClasse.Text = @"Traitement des statistiques";
        }

        private void ThreadStatistiquesTerminé(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 0;
            lblCompteur.Text = "";
            lblClasse.Text = @"Terminé !";
        }

        private void BtnSuppressionFichiers(object sender, EventArgs e)
        {
            string nomduFichierComposantes = "";
            string nomduFichierNotes = "";

            foreach (var fichierComposantes in chkLb_Composantes.CheckedItems)
            {
                foreach (DataRow ligne in TableComposantes.Rows)
                {
                    if (ligne[1].ToString() == fichierComposantes.ToString())
                    {
                        nomduFichierComposantes = ligne[0].ToString();
                    }
                }
                File.Delete(lblSource.Text + nomduFichierComposantes);
            }
            foreach (var fichierNotes in chkLb_Notes.CheckedItems)
            {
                foreach (DataRow ligne in TableNotes.Rows)
                {
                    if (ligne[1].ToString() == fichierNotes.ToString())
                    {
                        nomduFichierNotes = ligne[0].ToString();
                    }
                }
                File.Delete(lblSource.Text + nomduFichierNotes);
            }

            RemplirDatatable(TableNotes, lblSource.Text, "*.xls*", "Recapitulatif", "Notes", "AliasFichierNotes");
            RemplirDatatable(TableComposantes, lblSource.Text, "*.xls*", "Composantes", "Composantes", "AliasFichierComposantes");
            RemplirListeBox(chkLb_Notes, TableNotes);
            RemplirListeBox(chkLb_Composantes, TableComposantes);
        }

        private void CopieFichiersTypeDnb(Stream input, Stream output)
        {
            var buffer = new byte[32768];
            while (true)
            {
                var read = input.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                    return;
                output.Write(buffer, 0, read);
            }
        }

        private void RemplirDatatable(System.Data.DataTable dt, string folder, string fileType, string recherche, string nom, string alias)
        {
            dt.Rows.Clear();
            dt.Columns.Clear();
            var dinfo = new DirectoryInfo(folder);
            var files = dinfo.GetFiles(fileType).OrderByDescending(p => p.CreationTime);
            dt.Columns.Add(nom, typeof(String));
            dt.Columns.Add(alias, typeof(String));
            foreach (var file in files)
                if (file.Name.Contains(recherche))
                {
                    int index = file.Name.IndexOf('-');
                    string classe = file.Name.Substring(index - 1, 1);
                    dt.Rows.Add(file.Name, "3" + classe + " - " + nom + " au " + File.GetCreationTime(lblSource.Text + file.Name));
                }
        }

        private void RemplirListeBox(CheckedListBox lsb, System.Data.DataTable dt)
        {
            lsb.Items.Clear();
            foreach (DataRow ligne in dt.Rows)
                lsb.Items.Add(ligne[1]);
        }

        private void TuerProcessus(string processus)
        {
            var process = System.Diagnostics.Process.GetProcessesByName(processus);
            foreach (var p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
        }

        private int Bareme(int colonne)
        {
            int bareme = 0;
            if (colonne == 2) bareme = 50;
            if (colonne == 3) bareme = 40;
            if (colonne == 4) bareme = 10;
            if (colonne == 5) bareme = 50;
            if (colonne == 6) bareme = 100;
            if (colonne == 7) bareme = 25;
            if (colonne == 8) bareme = 25;
            return bareme;
        }

        private string NumDnb()
        {
            string numDnb = "";
            if (rdbDnb1.Checked == true) numDnb = "DNB1";
            if (rdbDnb2.Checked == true) numDnb = "DNB2";
            return numDnb;
        }

        private void chkLb_Notes_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable tableNotes = new System.Data.DataTable();
            System.Data.DataTable tableComposantes = new System.Data.DataTable();
            tableNotes.Columns.Add("notes", typeof(string));
            tableComposantes.Columns.Add("composantes", typeof(string));
            int i = 0;

            foreach (var item1 in chkLb_Notes.CheckedItems)
            {
                string classe1 = item1.ToString().Substring(0, 2);

                foreach (DataRow classe in tableNotes.Rows)
                {
                    foreach (var item in classe.ItemArray)
                    {
                        if (classe1 == item.ToString())
                            i = 1;
                    }
                }
                tableNotes.Rows.Add(classe1);
            }

            foreach (var item1 in chkLb_Composantes.CheckedItems)
            {
                string classe1 = item1.ToString().Substring(0, 2);

                foreach (DataRow classe in tableComposantes.Rows)
                {
                    foreach (var item in classe.ItemArray)
                    {
                        if (classe1 == item.ToString())
                            i = 1;
                    }
                }
                tableComposantes.Rows.Add(classe1);
            }

            DataTable dt;
            dt = GetDifferentRecords(tableNotes, tableComposantes);

            if ((dt.Rows.Count == 0) && (i == 0))
            {
                BtnGénérerDiplomes.Enabled = true;
                DataView dv = tableNotes.DefaultView;
                dv.Sort = "notes asc";
                DataTable tableNotes1 = dv.ToTable();
                foreach (DataRow classe in tableNotes1.Rows)
                {
                    foreach (var item in classe.ItemArray)
                    {
                        lblClasses.Text = lblClasses.Text + item + @"   ";
                    }
                }
            }
            else
            {
                BtnGénérerDiplomes.Enabled = false;
                lblClasses.Text = "";
            }
        }

        #region Compare two DataTables and return a DataTable with DifferentRecords

        public DataTable GetDifferentRecords(DataTable firstDataTable, DataTable secondDataTable)
        {
            //Create Empty Table
            DataTable resultDataTable = new DataTable("ResultDataTable");

            //use a Dataset to make use of a DataRelation object
            using (DataSet ds = new DataSet())
            {
                //Add tables
                ds.Tables.AddRange(new DataTable[] { firstDataTable.Copy(), secondDataTable.Copy() });

                //Get Columns for DataRelation
                DataColumn[] firstColumns = new DataColumn[ds.Tables[0].Columns.Count];
                for (int i = 0; i < firstColumns.Length; i++)
                {
                    firstColumns[i] = ds.Tables[0].Columns[i];
                }

                DataColumn[] secondColumns = new DataColumn[ds.Tables[1].Columns.Count];
                for (int i = 0; i < secondColumns.Length; i++)
                {
                    secondColumns[i] = ds.Tables[1].Columns[i];
                }

                //Create DataRelation
                DataRelation r1 = new DataRelation(string.Empty, firstColumns, secondColumns, false);
                ds.Relations.Add(r1);

                DataRelation r2 = new DataRelation(string.Empty, secondColumns, firstColumns, false);
                ds.Relations.Add(r2);

                //Create columns for return table
                for (int i = 0; i < firstDataTable.Columns.Count; i++)
                {
                    resultDataTable.Columns.Add(firstDataTable.Columns[i].ColumnName, firstDataTable.Columns[i].DataType);
                }

                //If FirstDataTable Row not in SecondDataTable, Add to ResultDataTable.
                resultDataTable.BeginLoadData();
                foreach (DataRow parentrow in ds.Tables[0].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r1);
                    if (childrows == null || childrows.Length == 0)
                        resultDataTable.LoadDataRow(parentrow.ItemArray, true);
                }

                //If SecondDataTable Row not in FirstDataTable, Add to ResultDataTable.
                foreach (DataRow parentrow in ds.Tables[1].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r2);
                    if (childrows == null || childrows.Length == 0)
                        resultDataTable.LoadDataRow(parentrow.ItemArray, true);
                }
                resultDataTable.EndLoadData();
            }

            return resultDataTable;
        }

        #endregion
    }
}