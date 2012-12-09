using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using itext = iTextSharp.text;
using ipdf = iTextSharp.text.pdf;
using System.Configuration;


namespace SQLToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        static SqlCommand myCommand;
        static SqlConnection myConnection;

        static DataTable dt = new DataTable();

        static string ExportFileLocation = @"c:\temp\";
        

        public MainWindow()
        {
            this.Title = "Easy SQL";
                   

            InitializeComponent();



            myConnection = new SqlConnection(SDN.LazyWork.Properties.Settings.Default.SQLConn);

            myConnection.Open();

            DataTable dt = ReturnSQLData("SELECT * FROM sys.Tables order by 1");
                        
            foreach (DataRow dr in dt.Rows)
	        {
                ddlTables.Items.Add(dr[0]);
	        }

            txtQuery.Text = "";

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            CreateExelFile();
            ExportToExcel();
            OpenAndSelectFile(".xlsx");
            
        }
               
      

        private void SubmitQuery(object sender, RoutedEventArgs e)
        {
            if(!txtQuery.Text.ToLower().StartsWith("sp_"))
                if (ddlTables.SelectedItem == null || txtQuery.Text.Length.Equals(0))
                {
                    MessageBox.Show("Tables or Query is blank.");
                    return;
                }
   
            string CmdString = txtQuery.Text;
            SqlCommand cmd = new SqlCommand(CmdString, myConnection);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            dt.Clear();
            dt = new DataTable();

            try
            {
                sda.Fill(dt);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            
            
            datatgridme.ItemsSource = dt.DefaultView;
            lblRowCount.Content = "Rows: " + dt.Rows.Count;
            

        }

        private void ddlTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataTable dt = CallSP(ddlTables.SelectedItem.ToString());
            string fields = string.Empty;

            foreach (DataRow dr in dt.Rows)
            {
                if(fields.Length.Equals(0))
                    fields = dr["column_name"].ToString();
                else
                    fields += ", " + dr["column_name"];
            }

            string Top = string.Empty;

            if(ckbTop.IsChecked == true)
                Top = " TOP 100 ";

            txtQuery.Text = string.Format("select {2} {1} from {0}",ddlTables.SelectedItem.ToString(),fields,Top);

            if(ckbAutoQuery.IsChecked.Equals(true))
                SubmitQuery(null, null);

        }

        private void Window_Closing_1(object sender, System.ComponentModel.CancelEventArgs e)
        {
            myConnection.Close();
            myConnection.Dispose();
            myCommand.Dispose();
            dt.Clear();
            dt.Dispose();
        }

        #region MyCode
        private DataTable CallSP(string TableName)
        {
            myCommand = new SqlCommand("sp_columns", myConnection);

            myCommand.CommandType = CommandType.StoredProcedure;

            myCommand.Parameters.Add(new SqlParameter("@table_name", TableName));

            SqlDataReader rdr = myCommand.ExecuteReader();

            DataTable dt = new DataTable();
            dt.Load(rdr);

            return dt;


        }

        public DataTable ReturnSQLData(string SQLStatment)
        {
            myCommand = new SqlCommand(SQLStatment, myConnection);

            SqlDataAdapter dscmd = new SqlDataAdapter(SQLStatment, myConnection);
            DataSet ds = new DataSet();
            dscmd.Fill(ds);

            return ds.Tables[0];



        }

        private void CreateExelFile()
        {
            string fileName = ExportFileLocation + ddlTables.SelectedItem.ToString() + ".xlsx";

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = ddlTables.SelectedItem.ToString()
            };
            sheets.Append(sheet);

            // Close the document.
            spreadsheetDocument.Close();

            //Console.WriteLine("The spreadsheet document has been created.\nPress a key.");

        }

        /// <summary>
        /// Using openxml
        /// </summary>
        private void ExportToExcel()
        {
            // Open the copied template workbook. 
            using (SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(ExportFileLocation + ddlTables.SelectedItem.ToString() + ".xlsx", true))
            {
                // Access the main Workbook part, which contains all references.
                WorkbookPart workbookPart = myWorkbook.WorkbookPart;

                // Get the first worksheet. 
                //WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(2);
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(0);

                // The SheetData object will contain all the data.
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Begining Row pointer                       
                int index = 2;

                Row row = new Row();
                row.RowIndex = (UInt32)1;

                #region Making headers


                for (int i = 0; i < dt.Columns.Count; i++)
                {

                    // New Cell
                    Cell cell = new Cell();
                    cell.DataType = CellValues.InlineString;
                    // Column A1, 2, 3 ... and so on
                    cell.CellReference = Convert.ToChar(65 + i).ToString() + "1";

                    // Create Text object
                    Text t = new Text();
                    t.Text = dt.Columns[i].ColumnName;

                    // Append Text to InlineString object
                    InlineString inlineString = new InlineString();
                    inlineString.AppendChild(t);

                    // Append InlineString to Cell
                    cell.AppendChild(inlineString);

                    // Append Cell to Row
                    row.AppendChild(cell);

                }
                // Append Row to SheetData
                sheetData.AppendChild(row);
                #endregion

                // For each item in the database, add a Row to SheetData.
                foreach (DataRow dr in dt.Rows)
                {
                    // New Row
                    row = new Row();
                    row.RowIndex = (UInt32)index;


                    for (int i = 0; i < dt.Columns.Count; i++)
                    {

                        // New Cell
                        Cell cell = new Cell();
                        cell.DataType = CellValues.InlineString;
                        // Column A1, 2, 3 ... and so on
                        cell.CellReference = Convert.ToChar(65 + i).ToString() + index;

                        // Create Text object
                        Text t = new Text();
                        t.Text = dr[i].ToString();

                        // Append Text to InlineString object
                        InlineString inlineString = new InlineString();
                        inlineString.AppendChild(t);

                        // Append InlineString to Cell
                        cell.AppendChild(inlineString);

                        // Append Cell to Row
                        row.AppendChild(cell);

                    }
                    // Append Row to SheetData
                    sheetData.AppendChild(row);
                    // increase row pointer
                    index++;
                }

                // save
                worksheetPart.Worksheet.Save();
                myWorkbook.Dispose();
            }
        }

        /// <summary>
        /// Using excel dll. The old way.
        /// </summary>
        private void ExportToExcelv2()
        {
            //DataTable dtMainSQLData = ReturnSQLData("select GUserID,Username,DeletedDateTime,LoginName from GUser where DeletedDateTime is null and Username not like 'deleted%'");

            DataView temp = (DataView)datatgridme.ItemsSource;

            DataTable dtMainSQLData = temp.Table;

            DataColumnCollection dcCollection = dtMainSQLData.Columns;
            // Export Data into EXCEL Sheet
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            for (int i = 1; i < dtMainSQLData.Rows.Count + 1; i++)
            {
                for (int j = 1; j < dtMainSQLData.Columns.Count + 1; j++)
                {
                    if (i == 1)
                        ExcelApp.Cells[i, j] = dcCollection[j - 1].ToString();
                    else
                        ExcelApp.Cells[i, j] = dtMainSQLData.Rows[i - 1][j - 1].ToString();
                }
            }
            ExcelApp.ActiveWorkbook.SaveCopyAs("C:\\Users\\snewsom\\Desktop\\test.xls");
            ExcelApp.ActiveWorkbook.Saved = true;
            ExcelApp.Quit();
        }

        private void OpenAndSelectFile(string ext)
        {
            Process p = new Process();
            p.StartInfo.FileName = "explorer.exe";
            p.StartInfo.Arguments = @"/select, " + ExportFileLocation + ddlTables.SelectedItem.ToString() + ext;

            p.Start();


        }

        #endregion

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            // step 1: creation of a document-object
            itext.Document document = new itext.Document();

            // step 2:
            // we create a writer that listens to the document
            // and directs a PDF-stream to a file
            ipdf.PdfWriter.GetInstance(document, new FileStream( ExportFileLocation + ddlTables.SelectedItem + ".pdf", FileMode.Create));

            // step 3: we open the document
            document.Open();

            // step 4: we add a paragraph to the document
            //document.Add(new itext.Paragraph("Hello World"));
            
            string col = "";
           for (int i = 0; i < dt.Columns.Count; i++)
                {

                    col += "     " + dt.Columns[i].ColumnName;

                    
                }
                // Append Row to SheetData
                document.Add(new itext.Paragraph(col));
                

                // For each item in the database, add a Row to SheetData.
                foreach (DataRow dr in dt.Rows)
                {
                
                    string row = "";
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {

                        row += "     " + dr[i].ToString();

                
                    }
                    document.Add(new itext.Paragraph(row));
                }

            

            // step 5: we close the document
            document.Close();

            OpenAndSelectFile(".pdf");

        }

        private void ckbTop_Checked(object sender, RoutedEventArgs e)
        {
            if (ddlTables.SelectedItem == null || txtQuery.Text.Length.Equals(0))
                return;
            else
                ddlTables_SelectionChanged(null, null);
        }



    }
}
