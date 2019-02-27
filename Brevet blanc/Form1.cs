using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

                for (int i = 2; i <= effectif + 1; i++) //Copie du nom des élèves et de la classe
                {
                    récapitulatif.Cells[i, 1].Value = worksheet.Cells[3, i].Value.ToString();
                    récapitulatif.Cells[i, 2].Value = Classe;
                    épreuvesEcrites.Cells[i, 1].Value = worksheet.Cells[3, i].Value.ToString();
                }

                for (int i = 2; i <= effectif + 1; i++) //Copie des points des composantes
                {
                    for (int j = 3; j <= 10; j++)
                    {
                        if (worksheet2.Cells[j + 1, i].Value != null)
                            récapitulatif.Cells[i, j].Value = worksheet2.Cells[j + 1, i].Value.ToString();
                    }
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
            RowCount = chkLb_Notes.CheckedItems.Count;
            k = 0;
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

                k++;
                ThreadDiplomes.ReportProgress(k);
            }

            #endregion CopieNotes

            TuerProcessus("Excel");

            #region PublipostageDNB

            var fichiersDnb = Directory.GetFiles(lblDestination.Text + @"DNB\", "*.*");
            RowCount = 0;
            k = 0;
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
            lblCompteur.Text = e.ProgressPercentage + @" / " + RowCount;
            lblClasse.Text = @"Traitement des " + Classe;
        }

        private void ThreadDiplomesTerminé(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 0;
            lblCompteur.Text = "";
            lblClasse.Text = @"Terminé !";
        }

        private void BtnGénérerStatistiques(object sender, EventArgs e)
        {
            var fichiersDnbXlsx = Directory.GetFiles(lblDestination.Text + @"DNB\Notes\");
            var fichierStat = lblDestination.Text + @"DNB\Statistiques.xlsx";
            if (!File.Exists(fichierStat)){
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

            #region Dnb1SynthèseEtListing

            int ligne = 3;
            int ligne1 = 3;
            int ligne2 = 3;
            int ligne3 = 3;
            int ligneEleve = 2;
            statSynthèse.Range["B4:G13"].Value = 0;
            statMoyennes.Range["B4:I13"].Value = 0;

            foreach (var file in fichiersDnbXlsx)
            {
                var fichierDnbXlsx = Path.GetFileName(file);
                if (fichierDnbXlsx.Contains("DNB1") && fichierDnbXlsx.Contains("xlsx"))
                {
                    ligne++;
                    var fichierDnb = lblDestination.Text + @"DNB\Notes\" + fichierDnbXlsx;
                    var dnbXlsx = excelApplication.Workbooks.Open(fichierDnb);
                    var dnbRécapitulatif = (Worksheet)dnbXlsx.Sheets.Item[1];
                    var épreuvesEcrites = (Worksheet)dnbXlsx.Sheets.Item[2];

                    var range = dnbRécapitulatif.Range["AG2:AG50"];
                    var colMoyennes = dnbRécapitulatif.Range["AF2:AF50"];

                    statSynthèse.Range["A" + ligne].Value = dnbRécapitulatif.Range["B2"].Value.ToString();

                    foreach (Range element in range.Cells)
                    {
                        if (element.Value2 != null)
                        {
                            statSynthèse.Range["B" + ligne].Value =
                                    int.Parse(statSynthèse.Range["B" + ligne].Value.ToString()) + 1;

                            if (element.Value.ToString().Contains("Non"))
                            {
                                statSynthèse.Range["C" + ligne].Value =
                                    int.Parse(statSynthèse.Range["C" + ligne].Value.ToString()) + 1;
                                statListing.Range["A" + ligne1].Value =
                                    dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + épreuvesEcrites.Range["A" + ligneEleve].Value.ToString() + " (" + dnbRécapitulatif.Range["AG" + ligneEleve].Value.ToString() + ")";
                                ligne1++;
                                int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value / 2 -
                                            dnbRécapitulatif.Range["AE" + ligneEleve].Value);
                                if ((delta <= numDelta.Value) && (ligne3 <= 31))
                                {
                                    statDelta.Range["A" + ligne3].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour obtention)";
                                    ligne3++;
                                }
                                if ((delta <= numDelta.Value) && (ligne3 > 31))
                                {
                                    statDelta.Range["E" + (ligne3-29)].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour obtention)";
                                    ligne3++;
                                }
                            }
                            if (element.Value.ToString().Contains("sans mention"))
                            {
                                statSynthèse.Range["D" + ligne].Value =
                                    int.Parse(statSynthèse.Range["D" + ligne].Value.ToString()) + 1;

                                int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value * 12/20 -
                                            dnbRécapitulatif.Range["AE" + ligneEleve].Value);
                                if ((delta <= numDelta.Value) && (ligne3 <= 31))
                                {
                                    statDelta.Range["A" + ligne3].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention AB)";
                                    ligne3++;
                                }
                                if ((delta <= numDelta.Value) && (ligne3 > 31))
                                {
                                    statDelta.Range["E" + (ligne3 - 29)].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention AB)";
                                    ligne3++;
                                }
                            }
                            if (element.Value.ToString().Contains("mention AB"))
                            {
                                statSynthèse.Range["E" + ligne].Value =
                                    int.Parse(statSynthèse.Range["E" + ligne].Value.ToString()) + 1;

                                int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value * 14 / 20 -
                                            dnbRécapitulatif.Range["AE" + ligneEleve].Value);
                                if ((delta <= numDelta.Value) && (ligne3 <= 31))
                                {
                                    statDelta.Range["A" + ligne3].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention B)";
                                    ligne3++;
                                }
                                if ((delta <= numDelta.Value) && (ligne3 > 31))
                                {
                                    statDelta.Range["E" + (ligne3 - 29)].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention B)";
                                    ligne3++;
                                }
                            }
                            if (element.Value.ToString().Contains("mention B"))
                            {
                                statSynthèse.Range["F" + ligne].Value =
                                    int.Parse(statSynthèse.Range["F" + ligne].Value.ToString()) + 1;

                                int delta = Convert.ToInt32(dnbRécapitulatif.Range["AR" + ligneEleve].Value * 16 / 20 -
                                            dnbRécapitulatif.Range["AE" + ligneEleve].Value);
                                if ((delta <= numDelta.Value) && (ligne3 <= 31))
                                {
                                    statDelta.Range["A" + ligne3].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention TB)";
                                    ligne3++;
                                }
                                if ((delta <= numDelta.Value) && (ligne3 > 31))
                                {
                                    statDelta.Range["E" + (ligne3 - 29)].Value =
                                        dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + "  (manque " + delta + " points pour mention TB)";
                                    ligne3++;
                                }
                            }
                            if (element.Value.ToString().Contains("mention TB"))
                            {
                                statSynthèse.Range["G" + ligne].Value =
                                    int.Parse(statSynthèse.Range["G" + ligne].Value.ToString()) + 1;
                                statListing.Range["E" + ligne2].Value =
                                    dnbRécapitulatif.Range["B" + ligneEleve].Value.ToString() + " - " + dnbRécapitulatif.Range["A" + ligneEleve].Value.ToString() + " (" + dnbRécapitulatif.Range["AE" + ligneEleve].Value.ToString() + " / " + dnbRécapitulatif.Range["AR" + ligneEleve].Value.ToString() + ")";
                                ligne2++;
                            }

                            statSynthèse.Range["H" + ligne].Formula = "=SUM(D" + ligne + ":G" + ligne + ")";
                            statSynthèse.Range["I" + ligne].Formula = "=H" + ligne + "/B" + ligne;
                        }
                        if (ligneEleve == 50) ligneEleve = 2;
                        else
                            ligneEleve++;
                    }
                    float total = 0;
                    int compteur = 0;
                    foreach (Range element in colMoyennes.Cells)
                    {
                        if (element.Value2 != null)
                        {
                            total = total + float.Parse(element.Value.ToString());
                            compteur++;
                        }
                    }
                    statSynthèse.Range["J" + ligne].Value = total / compteur;
                    dnbXlsx.Close();
                }
            }
            var range1 = statSynthèse.Range["A" + (ligne + 1), "I13"];
            range1.Value = "";
            statSynthèse.Range["A" + (ligne + 2)].Value = "Niveau";

            var colonne = 'B';
            for (int i = 1; i < 8; i++)
            {
                statSynthèse.Range[colonne.ToString() + (ligne + 2)].Formula = "=SUM(" + colonne + "4:" + colonne + ligne + ")";
                colonne++;
            }

            statSynthèse.Range["I" + (ligne + 2)].Formula = "=H" + (ligne + 2) + "/B" + (ligne + 2);
            statSynthèse.Range["J" + (ligne + 2)].Formula = "=AVERAGE(J4:J" + ligne;

            #endregion Dnb1SynthèseEtListing

            #region Dnb1MoyennesEpreuves

            ligne = 3;

            foreach (var file in fichiersDnbXlsx)
            {
                var fichierDnbXlsx = Path.GetFileName(file);

                if (fichierDnbXlsx.Contains("DNB1") && fichierDnbXlsx.Contains("xlsx"))
                {
                    ligne++;
                    var fichierDnb = lblDestination.Text + @"DNB\Notes\" + fichierDnbXlsx;
                    var dnbXlsx = excelApplication.Workbooks.Open(fichierDnb);
                    var dnbEpreuvesEcrites = (Worksheet)dnbXlsx.Sheets.Item[2];
                    statMoyennes.Range["A" + ligne].Value = ((Worksheet)dnbXlsx.Sheets.Item[1]).Range["B2"].Value.ToString(); //classe
                    statMoyennes.Range["I" + ligne].Value = ""; //oral
                    int effectif = int.Parse(statSynthèse.Range["B" + ligne.ToString()].Value.ToString()); //effectif
                    colonne = 'B';
                    for (int i = 1; i < 8; i++)
                    {
                        int barême = int.Parse(dnbEpreuvesEcrites.Range[colonne + "1"].Value.ToString().Split(new[] { '/', ')' })[1]);

                        dnbEpreuvesEcrites.Range[colonne.ToString() + (effectif + 3)].Formula = "=AVERAGE(" + colonne.ToString() + "2:" + colonne.ToString() + (effectif + 2) + ")";

                        statMoyennes.Range[colonne.ToString() + ligne].Value = Math.Round(
                            float.Parse(dnbEpreuvesEcrites.Range[colonne.ToString() + (effectif + 3)].Value.ToString()) / barême * 20, 2);

                        colonne++;
                    }
                    dnbEpreuvesEcrites.Range["J" + (effectif + 3)].Formula = "=AVERAGE(B" + (effectif + 3) + ":H" + (effectif + 3) + ")";
                    statMoyennes.Range["J" + ligne].Value = Math.Round(float.Parse(dnbEpreuvesEcrites.Range["J" + (effectif + 3)].Value.ToString()), 2);

                    object misValue = Missing.Value;
                    dnbXlsx.Close(false, misValue, misValue);
                }
            }

            var range3 = statMoyennes.Range["A" + (ligne + 1), "J13"];
            range3.Value = "";

            statMoyennes.Range["A1"].Value = "Année scolaire 2018-2019";
            statMoyennes.Range["A" + (ligne + 2)].Value = "Niveau";

            colonne = 'B';
            for (int i = 1; i < 8; i++)
            {
                statMoyennes.Range[colonne.ToString() + (ligne + 2)].Formula = "=AVERAGE(" + colonne + "4:" + colonne + ligne + ")";
                colonne++;
            }
            statMoyennes.Range["J" + (ligne + 2)].Formula = "=AVERAGE(J4:J" + ligne + ")";

            #endregion Dnb1MoyennesEpreuves

            #region Dnb1MoyennesControleContinu

            ligne = 3;

            foreach (var file in fichiersDnbXlsx)
            {
                var fichierDnbXlsx = Path.GetFileName(file);

                if (fichierDnbXlsx.Contains("DNB1") && fichierDnbXlsx.Contains("xlsx"))
                {
                    ligne++;
                    var colonne1 = 'C';
                    var fichierDnb = lblDestination.Text + @"DNB\Notes\" + fichierDnbXlsx;
                    var dnbXlsx = excelApplication.Workbooks.Open(fichierDnb);
                    //var dnbEpreuvesEcrites = (Worksheet)dnbXlsx.Sheets.Item[2];
                    var dnbRécapitulatif = (Worksheet)dnbXlsx.Sheets.Item[1];
                    statMoyennesControle.Range["A" + ligne].Value = ((Worksheet)dnbXlsx.Sheets.Item[1]).Range["B2"].Value.ToString();
                    statMoyennesControle.Range["I" + ligne].Value = "";
                    int effectif = int.Parse(statSynthèse.Range["B" + ligne.ToString()].Value.ToString());
                    colonne = 'B';
                    for (int i = 1; i < 9; i++)
                    {
                        float somme = 0;
                        for (int j = 2; j <= effectif + 1; j++)
                        {
                            if (dnbRécapitulatif.Range[colonne1.ToString() + j].Value != null)
                                somme = somme + float.Parse(dnbRécapitulatif.Range[colonne1.ToString() + j].Value.ToString());
                        }

                        statMoyennesControle.Range[colonne.ToString() + ligne].Value = somme / effectif;

                        colonne++;
                        colonne1++;
                    }

                    statMoyennesControle.Range["J" + ligne].Formula = "=AVERAGE(B" + ligne + ":I" + ligne;
                    dnbXlsx.Close();
                }
            }

            var range5 = statMoyennesControle.Range["A" + (ligne + 1), "J13"];
            range5.Value = "";

            statMoyennesControle.Range["A1"].Value = "Année scolaire 2018-2019";
            statMoyennesControle.Range["A" + (ligne + 2)].Value = "Niveau";

            colonne = 'B';
            for (int i = 1; i < 9; i++)
            {
                statMoyennesControle.Range[colonne.ToString() + (ligne + 2)].Formula = "=AVERAGE(" + colonne + "4:" + colonne + ligne + ")";
                colonne++;
            }
            statMoyennesControle.Range["J" + (ligne + 2)].Formula = "=AVERAGE(J4:J" + ligne + ")";

            #endregion Dnb1MoyennesControleContinu

            excelApplication.DisplayAlerts = false;
            statXlsx.SaveAs(fichierStat);
            statXlsx.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                            lblDestination.Text + @"DNB\Statistiques.pdf");
            statXlsx.Close();
            GC.Collect();
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
    }
}