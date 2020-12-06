/*
 * Created by SharpDevelop.
 * User: jlicciardello
 * Date: 20/11/2020
 * Time: 11:26
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using System.Net;
using System.IO	;

using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.SpreadSheet;

namespace identify_abbreviations
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		String ris = "";
		void Button1Click(object sender, EventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Filter = "HTML|*.html";
			
			if(ofd.ShowDialog() == DialogResult.OK)
			{
				richTextBox1.Clear();
				//INIZIALIZZAZIONE VARIABILI PER FILTRI CON CHECKBOX
				string puntoSpazio = "";
				string spazioPunto = "";
				string puntoSingolo = "";
				if (checkBox1.Checked) {
					puntoSpazio = ". ";
				}
				if (checkBox2.Checked) {
					spazioPunto = " .";
				}
				if (checkBox3.Checked) {
					puntoSingolo = ".";
				}
				//--------------------------------------------------

				string filePath = ofd.FileName;			//OTTIENE LA STRINGA CORRISPONDENTE AL PERCORSO DEL FILE SELEZIONATO
				//----------CODICE-PER-PRENDERE-IL-SORGENTE-DELL'HTML-------------
				FileWebRequest request=(FileWebRequest)WebRequest.Create(filePath);
				FileWebResponse response=(FileWebResponse)request.GetResponse();
				StreamReader sr=new StreamReader(response.GetResponseStream());	
				String htmlSource = sr.ReadToEnd();
				//----------------------------------------------------------------
				//Dentro htmlSource c'è una il sorgente

				//string[] testSplit = htmlSource.Split("<tr id=\"r1\">");
				string[] htmlSplitted = null;
				htmlSplitted = htmlSource.Split(new[] { "<tr id=" }, StringSplitOptions.None);



				//richTextBox1.Text = htmlSplitted2[0];

				double t = 100/htmlSplitted.Length;

				for(int i=1; i<htmlSplitted.Length; i++){
					
					string[] htmlSplitted1 = null;
					htmlSplitted1 =	htmlSplitted[i].Split(new[] { "<td class=\"td2\">" }, StringSplitOptions.None);
					//--------------------------------MODIFICHE PER AGGIUNGERE ID E SOURCE-------------------------------------------------
					string[] htmlSplitSup = null;                                                                   //Variabile che creo per prendere anche id source
					htmlSplitSup = htmlSplitted1[0].Split(new[] { "</td>" }, StringSplitOptions.None);

					string[] htmlSplitSup2 = null;
					htmlSplitSup2 = htmlSplitSup[0].Split(new[] { ">" }, StringSplitOptions.None);

					string[] htmlSplitSup3 = null;
					htmlSplitSup3 = htmlSplitSup[1].Split(new[] { "<td>" }, StringSplitOptions.None);

					string id = htmlSplitSup2[2];
					string source = htmlSplitSup3[1];
					//-----------------------------------------------------------------------------------------------------------------
					if (htmlSplitted1.Length<2){
					continue;
					}
					//LA STRINGA DI TARGHE E' DENTRO htmlSplitted2
					string[] htmlSplitted2 = null;
					htmlSplitted2 = htmlSplitted1[1].Split(new[] { "</td>" }, StringSplitOptions.None);
					
					
					if(htmlSplitted2[0].Contains(puntoSpazio) || htmlSplitted2[0].Contains(spazioPunto) || htmlSplitted2[0].Contains(puntoSingolo)){
						richTextBox1.Text += id + " | " + source + " | " + htmlSplitted2[0] +  "\n";
					}
					
				}
				//richTextBox1.Text = ris;
				sr.Close();
			}
		}
	
		
		void Button2Click(object sender, EventArgs e)
		{
			SaveFileDialog sfd = new SaveFileDialog();
			//sfd.InitialDirectory = @"C:\";
			sfd.RestoreDirectory = true;
			sfd.FileName = "*.xls";
			sfd.DefaultExt = "Parole abbreviate";
			sfd.Filter = "Xls files (*.xls)|*.xls";
			
			if(sfd.ShowDialog() == DialogResult.OK)
			{
			string file = sfd.FileName;
			Workbook workbook = new Workbook(); 
			Worksheet worksheet = new Worksheet("Parole abbreviate"); 
			worksheet.Cells[0, 0] = new Cell("ID");
			//-----------MODIFICHE PER OTTENERE ID E SOURCE------------------------------------
			worksheet.Cells[0, 1] = new Cell("SOURCE");
			worksheet.Cells[0, 2] = new Cell("TARGHET");

			string[] tempSplit = ris.Split(new[] { "\n" }, StringSplitOptions.None);
			

				for (int i=0; i<tempSplit.Length; i++){
					//------------------------ROBA AGGIUNTA-------------------------------------------------
					string[] tempSplit1= null;
					tempSplit1 = tempSplit[i].Split(new[] { " | " }, StringSplitOptions.None);
					if (tempSplit1.Length<3) {
						continue;
					}
					worksheet.Cells[i + 1, 0] = new Cell(tempSplit1[0]);
					worksheet.Cells[i + 1, 1] = new Cell(tempSplit1[1]);
					worksheet.Cells[i + 1, 2] = new Cell(tempSplit1[2]);
					//----------------------------------------------------------------------------------------
				}
			
			workbook.Worksheets.Add(worksheet);
			workbook.Save(file);
			MessageBox.Show("File excel salvato con successo");
			}
			


		}

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
