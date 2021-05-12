using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Aspose.Cells;
using System.Threading;


namespace Задание
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			GetClients();
		}
		private void GetClients()
		{
			SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-16HQO5L\MSSQLSERVER01;Initial Catalog=KeremetBank;Integrated Security=True");
			SqlCommand cmd = new SqlCommand("select*from Clients", con);
			DataTable dt = new DataTable();
			con.Open();
			SqlDataReader sdr = cmd.ExecuteReader();
			dt.Load(sdr);
			con.Close();
			dataGridView1.DataSource = dt;
		}

		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			//dataGridView1.CurrentRow.Selected = true;
			IDTextBox.Text = dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString();
			NameTextBox.Text = dataGridView1.Rows[e.RowIndex].Cells["Name"].Value.ToString();
			BirthDatedateTimePicker.Text = dataGridView1.Rows[e.RowIndex].Cells["BirthDate"].Value.ToString();
			PhoneNumberTextBox.Text = dataGridView1.Rows[e.RowIndex].Cells["PhoneNumber"].Value.ToString();
			AddressTextBox.Text = dataGridView1.Rows[e.RowIndex].Cells["Address"].Value.ToString();
			SocialNumberTextBox.Text = dataGridView1.Rows[e.RowIndex].Cells["SocialNumber"].Value.ToString();
		}

		private void PrintButton_Click(object sender, EventArgs e)
		{
			if (IDTextBox.Text == "")
			{
				MessageBox.Show("Select client");

			}
			else
			{
				
								  // Load Excel workbook
				Workbook workbook = new Workbook("C:\\Users\\Azat\\source\\repos\\Задание\\Template\\example.xlsx");
				ReplaceOptions replace = new ReplaceOptions();
				// Set case sensitivity and text matching options
				replace.CaseSensitive = false;
				replace.MatchEntireCellContents = false;
				// Replace text
				workbook.Replace("[Date]", DateTime.Now.ToString(), replace);
				workbook.Replace("[ID]", IDTextBox.Text, replace);
				workbook.Replace("[Name]", NameTextBox.Text, replace);
				workbook.Replace("[BirthDate]", BirthDatedateTimePicker.Text, replace);
				workbook.Replace("[PhoneNumber]", PhoneNumberTextBox.Text, replace);
				workbook.Replace("[Address]", AddressTextBox.Text, replace);
				workbook.Replace("[SocialNumber]", SocialNumberTextBox.Text, replace);
				// Save updated Excel workbook (file name=date)
				string filename = DateTime.Now.ToString("dd.MM.yyyy-HH.mm");
				workbook.Save("C:\\Users\\Azat\\source\\repos\\Задание\\Result\\" + filename + ".xlsx") ;



				MessageBox.Show("Export to Excel is successful");
			}
		}
	}
}
