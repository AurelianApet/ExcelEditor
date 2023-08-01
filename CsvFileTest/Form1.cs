//////////////////////////////////////////////////////////////////////////////
// This source code and all associated files and resources are copyrighted by
// the author(s). This source code and all associated files and resources may
// be used as long as they are used according to the terms and conditions set
// forth in The Code Project Open License (CPOL).
//
// Copyright (c) 2012 Jonathan Wood
// http://www.blackbeltcoder.com
//

using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using CsvFile;
using System.Drawing;
using System.Data;
using System.ComponentModel;
using Microsoft.Win32;
using Excel;
using ICSharpCode;
using ICSharpCode.SharpZipLib;
using System.Linq;
using System.Net.Mail;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MyExcel = Microsoft.Office.Interop.Excel;
namespace CsvFileTest
{
    public partial class Form1 : Form
    {
        private int MaxColumns = 20;
        protected string FileName;
        protected bool Modified;
        private MyExcel.Workbook workbook;
        private MyExcel.Worksheet worksheet;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitializeGrid();
            ClearFile();
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SaveIfModified())
                ClearFile();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SaveIfModified())
            {
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";
                if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
                    ReadFile(openFileDialog1.FileName);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (FileName != null)
                WriteFile(FileName);
            else
                saveAsToolStripMenuItem_Click(sender, e);
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = FileName;
            if (saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                string ext = Path.GetExtension(saveFileDialog1.FileName);
                if(ext == ".csv")
                {
                    WriteFile(saveFileDialog1.FileName);
                }
                else
                {
                    excelsave(saveFileDialog1.FileName, workbook, worksheet);
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SaveIfModified())
                Close();
        }

        /// <summary>
        /// //////////////////////////////////////////////
        /// </summary>

        private void InitializeGrid()
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear(); ;
            for (int i = 1; i <= MaxColumns; i++)
            {
                dataGridView1.Columns.Add(
                    String.Format("Column{0}", i),
                    String.Format("Column {0}", i));
            }
        }

        private void ClearFile()
        {
            dataGridView1.Rows.Clear();
            FileName = null;
            Modified = false;
        }

        private bool ReadFile(string filename)
        {
            Cursor = Cursors.WaitCursor;
            string ext = Path.GetExtension(filename);
            if(ext == ".csv")
            {
                try
                {
                    dataGridView1.Rows.Clear();
                    List<string> columns = new List<string>();
                    using (var reader = new CsvFileReader(filename))
                    {
                        while (reader.ReadRow(columns))
                        {
                            dataGridView1.Rows.Add(columns.ToArray());
                        }
                    }
                    FileName = filename;
                    Modified = false;
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(String.Format("Error reading from {0}.\r\n\r\n{1}", filename, ex.Message));
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
            else
            {
                MyExcel.Application app = new Microsoft.Office.Interop.Excel.Application();
                try
                {
                    workbook = app.Workbooks.Open(filename);
                    worksheet = workbook.ActiveSheet;
                    int rcount = worksheet.UsedRange.Rows.Count;
                    int ccount = worksheet.UsedRange.Columns.Count;
                    int i = 1;
                    for (int j = 0; j < ccount; j++)
                    {
                        dataGridView1.Columns[j].HeaderText = worksheet.Cells[1, j + 1].Value;
                    }

                    for (; i < rcount; i++)
                    {
                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                        for (int j = 0; j < ccount; j++)
                        {
                            row.Cells[j].Value = worksheet.Cells[i + 1, j + 1].Value;
                        }
                        dataGridView1.Rows.Add(row);
                    }
                }
                finally
                {
                    app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                    app = null;
                    Cursor = Cursors.Default;
                }
            }
            return false;
        }

        private bool WriteFile(string filename)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                // Like Excel, we'll get the highest column number used,
                // and then write out that many columns for every row
                int numColumns = GetMaxColumnUsed();
                using (var writer = new CsvFileWriter(filename))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            List<string> columns = new List<string>();
                            for (int col = 0; col < numColumns; col++)
                                columns.Add((string)row.Cells[col].Value ?? String.Empty);
                            writer.WriteRow(columns);
                        }
                    }
                }
                FileName = filename;
                Modified = false;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Error writing to {0}.\r\n\r\n{1}", filename, ex.Message));
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            return false;
        }

        // Determines the maximum column number used in the grid
        private int GetMaxColumnUsed()
        {
            int maxColumnUsed = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    for (int col = row.Cells.Count - 1; col >= 0; col--)
                    {
                        if (row.Cells[col].Value != null)
                        {
                            if (maxColumnUsed < (col + 1))
                                maxColumnUsed = (col + 1);
                            continue;
                        }
                    }
                }
            }
            return maxColumnUsed;
        }

        private bool SaveIfModified()
        {
            if (!Modified)
                return true;

            DialogResult result = MessageBox.Show("The current file has changed. Save changes?", "Save Changes", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            {
                if (FileName != null)
                {
                    return WriteFile(FileName);
                }
                else
                {
                    saveFileDialog1.FileName = FileName;
                    if (saveFileDialog1.ShowDialog(this) == DialogResult.OK)
                    {
                        string ext = Path.GetExtension(FileName);
                        if (ext == ".csv")
                        {
                            return WriteFile(saveFileDialog1.FileName);
                        }
                        else
                        {
                            excelsave(FileName, workbook, worksheet);
                        }
                    }
                    return false;
                }
            }
            else if (result == DialogResult.No)
            {
                return true;
            }
            else // DialogResult.Cancel
            {
                return false;
            }
        }

        static void excelsave(string filename, MyExcel.Workbook xlWorkBook, MyExcel.Worksheet xlWorkSheet)
        {
            MyExcel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            try
            {
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook.SaveAs(filename, MyExcel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue,
                    false, false, MyExcel.XlSaveAsAccessMode.xlNoChange, MyExcel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
                //            xlWorkBook.SaveAs(filename, MyExcel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, MyExcel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                MessageBox.Show("Excel file created successfully!");
            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }
            xlApp.Quit();
        }

        public void DataTableToExcel()
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                // creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet
                worksheet.Name = "Sheet1";
                // storing header part in Excel
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                // storing Each row and column value to excel sheet
                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        string values = string.Empty;
                        values = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        worksheet.Cells[i + 2, j + 1] = values;
                    }
                }
            }
            finally
            {
                //Release the resources
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }
        }

        public static string ShowDialog(string caption)
        {
            Form prompt = new Form();
            prompt.Width = 450;
            prompt.Height = 250;
            prompt.Text = caption;
            Button confirmation = new Button() { Text = "Input", Left = 30, Width = 90, Top = 100 };
            TextBox inputBox = new TextBox() { Left = 150, Top = 100, Width = 250 };
            confirmation.Click += (sender, e) =>
            {
                prompt.Close();
            };
            inputBox.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    prompt.Close();
                }
            };
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(inputBox);
            prompt.Shown += (s, e) => inputBox.Focus();
            prompt.ShowDialog();
            return inputBox.Text;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Modified = true;
        }

        private void insertColumnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                int columnsnum = Convert.ToInt16(ShowDialog("Input the numbers of the insert columns"));

                int columnindex = -1;
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int columnindex1 = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    if (columnindex1 > columnindex) columnindex = columnindex1;
                }
                DataGridViewTextBoxColumn txtColum = new DataGridViewTextBoxColumn();
                txtColum.HeaderText = "Column" + Convert.ToString(columnindex);
                for(int i = 0; i < columnsnum + 1; i ++)
                {
                    dataGridView1.Columns.Insert(columnindex + i, txtColum);
                }
                //                MaxColumns += columnsnum;
                //                InitializeGrid();
            }
            catch(Exception ex)
            {

            }
        }

        private void deleteColumnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int columnindex1 = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    dataGridView1.Columns.RemoveAt(columnindex1);
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void insertRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int rowsnum = Convert.ToInt16(ShowDialog("Input the numbers of the insert rows"));
                Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                int rowindex = -1;
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex1 = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    if (rowindex1 > rowindex) rowindex = rowindex1;
                }
                dataGridView1.Rows.Insert(rowindex + 1, rowsnum);
            }
            catch (Exception ex)
            {

            }
        }

        private void deleteRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex1 = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    dataGridView1.Rows.RemoveAt(rowindex1);
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void formatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Int32 selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
            string str = "";
            try
            {
                str = ShowFontDialog("Input the fontsize");
            }
            catch(Exception ex)
            {

            }
            string[] list = System.Text.RegularExpressions.Regex.Split(str, ",");
            int fontsize;
            Color forecolor, backcolor;
            string fontstyle, fonttype, border;
            try
            {
                fontsize = Convert.ToInt32(list[0]);
            }
            catch(Exception ex)
            {
                fontsize = 0;
            }
            try
            {
                forecolor = Color.FromArgb(Convert.ToInt32(list[1]));
            }
            catch(Exception ex)
            {
                forecolor = Color.Black;
            }
            try
            {
                fontstyle = list[2];
            }
            catch(Exception ex)
            {
                fontstyle = "";
            }
            try
            {
                backcolor = Color.FromArgb(Convert.ToInt32(list[3]));
            }
            catch (Exception ex)
            {
                backcolor = Color.White;
            }
            try
            {
                fonttype = list[4];
            }
            catch (Exception ex)
            {
                fonttype = "";
            }
            try
            {
                border = list[5];
            }
            catch(Exception ex)
            {               
                border = "";
            }

            if (border == "none")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            else if (border == "Raised")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Raised;
            }
            else if (border == "Raised Horizontal")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.RaisedHorizontal;
            }
            else if (border == "Raised Vertical")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.RaisedVertical;
            }
            else if (border == "Single")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            }
            else if (border == "Single Horizontal")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            }
            else if (border == "Single Vertical")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical;
            }
            else if (border == "Sunken")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Sunken;
            }
            else if (border == "Sunken Horizontal")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SunkenHorizontal;
            }
            else if (border == "Sunken Vertical")
            {
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SunkenVertical;
            }

            if (fontstyle == "Bold")
            {
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    int columnindex = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.Font = new System.Drawing.Font(fonttype, fontsize, FontStyle.Bold);
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.ForeColor = forecolor;
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.BackColor = backcolor;
                }
            }
            else if (fontstyle == "Italic")
            {
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    int columnindex = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.Font = new System.Drawing.Font(fonttype, fontsize, FontStyle.Italic);
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.ForeColor = forecolor;
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.BackColor = backcolor;
                }
            }
            else if (fontstyle == "Underline")
            {
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    int columnindex = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.Font = new System.Drawing.Font(fonttype, fontsize, FontStyle.Underline);
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.ForeColor = forecolor;
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.BackColor = backcolor;
                }
            }
            else if (fontstyle == "Regualr")
            {
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    int columnindex = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.Font = new System.Drawing.Font(fonttype, fontsize, FontStyle.Regular);
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.ForeColor = forecolor;
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.BackColor = backcolor;
                }
            }
            else if (fontstyle == "Strikeout")
            {
                for (int i = 0; i < selectedCellCount; i++)
                {
                    int rowindex = Convert.ToInt32(dataGridView1.SelectedCells[i].RowIndex.ToString());
                    int columnindex = Convert.ToInt32(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.Font = new System.Drawing.Font(fonttype, fontsize, FontStyle.Strikeout);
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.ForeColor = forecolor;
                    dataGridView1.Rows[rowindex].Cells[columnindex].Style.BackColor = backcolor;
                }
            }
        }

        public static string ShowFontDialog(string caption)
        {
            Form prompt = new Form();
            prompt.Width = 370;
            prompt.Height = 230;
            prompt.Text = caption;
            Label fontsizetext = new Label() { Text = "Fontsize", Left = 30, Width = 50, Top = 32 };
            ComboBox fontsize = new ComboBox() {Left = 80, Width = 50, Top = 30};
            fontsize.Items.Add( "1");            fontsize.Items.Add("2");            fontsize.Items.Add("3");            fontsize.Items.Add("4");            fontsize.Items.Add("5");
            fontsize.Items.Add("6");            fontsize.Items.Add("7");            fontsize.Items.Add("8");            fontsize.Items.Add("9");            fontsize.Items.Add("10");
            fontsize.Items.Add("11");            fontsize.Items.Add("12");            fontsize.Items.Add("13");            fontsize.Items.Add("14");            fontsize.Items.Add("15");
            fontsize.Items.Add("16");            fontsize.Items.Add("17");            fontsize.Items.Add("18");            fontsize.Items.Add("19");            fontsize.Items.Add("20");
            fontsize.Items.Add("21");            fontsize.Items.Add("22");            fontsize.Items.Add("23");            fontsize.Items.Add("24");            fontsize.Items.Add("25");
            fontsize.SelectedIndex = 9;
            Label fontcolortext = new Label() { Text = "Fontcolor:", Left = 170, Width = 60, Top = 32 };
            Label fontcolor = new Label() { Text = "select color", Left = 230, Width = 160, Top = 32 };

            Label fontstyletxt = new Label() { Text = "FontStyle:", Left = 30, Width = 60, Top = 62};
            ComboBox fontstyle = new ComboBox() { Left = 90, Width = 80, Top = 60 };
            fontstyle.Items.Add("Regular");
            fontstyle.Items.Add("Bold");
            fontstyle.Items.Add("Italic");
            fontstyle.Items.Add("Strikeout");
            fontstyle.Items.Add("Underline");
            fontstyle.SelectedIndex = 0;

            Label backcolortxt = new Label() { Text = "Backcolor", Left = 170, Width = 60, Top = 60};
            Label backcolor = new Label() { Text = "select color", Left = 230, Width = 160, Top = 60};

            Label fonttypetxt = new Label() { Text = "FontType", Left = 30, Width = 60, Top = 92 };
            ComboBox fonttype = new ComboBox() { Left = 90, Width = 150, Top = 90};
            fonttype.Items.Add("Agency FB"); fonttype.Items.Add("ALGERIAN"); fonttype.Items.Add("Ami R"); fonttype.Items.Add("Arial"); fonttype.Items.Add("Arial Rounded MT");
            fonttype.Items.Add("Arial Unicode MS"); fonttype.Items.Add("Baskerville Old Face"); fonttype.Items.Add("Batang"); fonttype.Items.Add("BatangChe"); fonttype.Items.Add("Bell MT");
            fonttype.Items.Add("Berlin Sans FB"); fonttype.Items.Add("Bernard MT"); fonttype.Items.Add("Bodoni MT"); fonttype.Items.Add("Book Antiqua"); fonttype.Items.Add("Bookman Old Style");
            fonttype.Items.Add("Bookshelf Symbol 7"); fonttype.Items.Add("Brltannlc"); fonttype.Items.Add("Broadway"); fonttype.Items.Add("Calibri"); fonttype.Items.Add("Californian FB");
            fonttype.Items.Add("Calisto MT"); fonttype.Items.Add("Cambria"); fonttype.Items.Add("Cambria Math"); fonttype.Items.Add("Candara"); fonttype.Items.Add("CASTELLAR");
            fonttype.Items.Add("Centaur"); fonttype.Items.Add("Century"); fonttype.Items.Add("Century Gothic"); fonttype.Items.Add("Century Schoolbook"); fonttype.Items.Add("Colonna MT");
            fonttype.Items.Add("Comic Sans MS"); fonttype.Items.Add("Consolas"); fonttype.Items.Add("Constantia"); fonttype.Items.Add("Cooper"); fonttype.Items.Add("COPPERPLATE GOTHIC");
            fonttype.Items.Add("Corbel"); fonttype.Items.Add("Courier"); fonttype.Items.Add("Courier New"); fonttype.Items.Add("Dotum"); fonttype.Items.Add("DotumChe");
            fonttype.Items.Add("ENGARVERS MT"); fonttype.Items.Add("Eras ITC"); fonttype.Items.Add("Expo M"); fonttype.Items.Add("FangSong"); fonttype.Items.Add("FELIX TITLING");
            fonttype.Items.Add("Fixedsys"); fonttype.Items.Add("Footlight MT"); fonttype.Items.Add("Franklin Gothic"); fonttype.Items.Add("Franklin Gothic Book"); fonttype.Items.Add("Georgia");
            fonttype.Items.Add("Gill Sans"); fonttype.Items.Add("Gill Snas MT"); fonttype.Items.Add("Gloucester MT"); fonttype.Items.Add("Goudy Old Style"); fonttype.Items.Add("Gulim");
            fonttype.Items.Add("GulimChe"); fonttype.Items.Add("Gungsuh"); fonttype.Items.Add("GungsubChe"); fonttype.Items.Add("Harrington"); fonttype.Items.Add("Headline R");
            fonttype.Items.Add("High Tower Text"); fonttype.Items.Add("HYGothic"); fonttype.Items.Add("HYGothic - Extra"); fonttype.Items.Add("HYGraphic"); fonttype.Items.Add("HYGungSo");
            fonttype.Items.Add("HYHeadLine"); fonttype.Items.Add("HYPMokGak"); fonttype.Items.Add("HYPost"); fonttype.Items.Add("HYSinMyeongJo"); fonttype.Items.Add("Impact");
            fonttype.Items.Add("Imprint MT Shadow"); fonttype.Items.Add("KaiTi"); fonttype.Items.Add("Kristen ITC"); fonttype.Items.Add("Latin"); fonttype.Items.Add("Lucida Bright");
            fonttype.Items.Add("Lucida Calligraphy"); fonttype.Items.Add("Lucida Console"); fonttype.Items.Add("Lucida Console"); fonttype.Items.Add("Lucida Fax"); fonttype.Items.Add("Lucida Sans");
            fonttype.Items.Add("Lucida Sans Unicode"); fonttype.Items.Add("Magic R"); fonttype.Items.Add("Maiandra GD"); fonttype.Items.Add("Malgun Gothic"); fonttype.Items.Add("Microsoft Sans Serif");
            fonttype.Items.Add("Microsoft YaHei"); fonttype.Items.Add("Microsoft YaHei UI"); fonttype.Items.Add("Modern"); fonttype.Items.Add("MoeumT R"); fonttype.Items.Add("MS Outlook");
            fonttype.Items.Add("MS Reference Specialty"); fonttype.Items.Add("MS Sans Serif"); fonttype.Items.Add("MS Serif"); fonttype.Items.Add("MT Extra"); fonttype.Items.Add("New Gulim");
            fonttype.Items.Add("Nina"); fonttype.Items.Add("NSimSun"); fonttype.Items.Add("OCR A"); fonttype.Items.Add("Palatino Linotype"); fonttype.Items.Add("Papyrus");
            fonttype.Items.Add("Perpetua"); fonttype.Items.Add("Poor Richard"); fonttype.Items.Add("Pyunji R"); fonttype.Items.Add("Rockwell"); fonttype.Items.Add("Roman");
            fonttype.Items.Add("Segoe"); fonttype.Items.Add("Segoe Marker"); fonttype.Items.Add("Segoe Print"); fonttype.Items.Add("Segoe Script"); fonttype.Items.Add("Segoe UI");
            fonttype.Items.Add("Segoe UI Emoji"); fonttype.Items.Add("Segoe UI Symbol"); fonttype.Items.Add("SHOWCARD GOTHIC"); fonttype.Items.Add("SimHei"); fonttype.Items.Add("SimSun");
            fonttype.Items.Add("SimSun - ExtB"); fonttype.Items.Add("Sitka Banner"); fonttype.Items.Add("Sitka Display"); fonttype.Items.Add("Sitka Heading"); fonttype.Items.Add("Sitka Small");
            fonttype.Items.Add("Sitka Subheading"); fonttype.Items.Add("Sitka Text"); fonttype.Items.Add("Small Fontr"); fonttype.Items.Add("STENCIL"); fonttype.Items.Add("Symbol");
            fonttype.Items.Add("System"); fonttype.Items.Add("Tahoma"); fonttype.Items.Add("TeamViewer12"); fonttype.Items.Add("Terminal"); fonttype.Items.Add("Times New Roman");
            fonttype.Items.Add("Trebuchet MS"); fonttype.Items.Add("Verdana"); fonttype.Items.Add("Weddings"); fonttype.Items.Add("Wingdings"); fonttype.Items.Add("Wingdings 2");
            fonttype.Items.Add("Wingdings 3"); fonttype.Items.Add("Yet R");
            fonttype.SelectedIndex = 10;

            Label cellbordertxt = new Label() { Text = "Cellborder", Left = 30, Width = 60, Top = 122};
            Button confirmation = new Button() { Text = "Input", Left = 100, Width = 90, Top = 150 };
            ComboBox cellborder = new ComboBox() { Left = 90, Width = 150, Top = 120};
            cellborder.Items.Add("None"); cellborder.Items.Add("Raised"); cellborder.Items.Add("Raised Horizontal");
            cellborder.Items.Add("Raised Vertical"); cellborder.Items.Add("Single"); cellborder.Items.Add("Single Horizontal"); cellborder.Items.Add("Single Vertical");
            cellborder.Items.Add("Sunken"); cellborder.Items.Add("Sunken Horizontal"); cellborder.Items.Add("Sunken Vertical");

            cellborder.SelectedIndex = 0;

            //            TextBox inputBox = new TextBox() { Left = 150, Top = 100, Width = 250 };
            confirmation.Click += (sender, e) =>
            {
                prompt.Close();
            };
//            inputBox.KeyDown += (sender, e) =>
//            {
//                if (e.KeyCode == Keys.Enter)
//                {
//                    prompt.Close();
//                }
//            };
            fontcolor.Click += (sender, e) =>
            {
                ColorDialog colorDialog1 = new ColorDialog();
                if (colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    fontcolor.Text = Convert.ToString(colorDialog1.Color.ToArgb().ToString());
                }
            };
            backcolor.Click += (sender, e) =>
            {
                ColorDialog colorDialog1 = new ColorDialog();
                if (colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    backcolor.Text = Convert.ToString(colorDialog1.Color.ToArgb().ToString());
                }
            };

            prompt.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    prompt.Close();
                }
            };
            prompt.Controls.Add(fontsizetext);
            prompt.Controls.Add(fontsize);
            prompt.Controls.Add(fontcolortext);
            prompt.Controls.Add(fontcolor);
            prompt.Controls.Add(fontstyletxt);
            prompt.Controls.Add(fontstyle);
            prompt.Controls.Add(backcolortxt);
            prompt.Controls.Add(backcolor);
            prompt.Controls.Add(fonttypetxt);
            prompt.Controls.Add(fonttype);
            prompt.Controls.Add(cellbordertxt);
            prompt.Controls.Add(cellborder);
            //            prompt.Controls.Add(inputBox);
            prompt.Controls.Add(confirmation);
//            prompt.Shown += (s, e) => inputBox.Focus();
            prompt.ShowDialog();
            string str = fontsize.Text + "," + fontcolor.Text + "," + fontstyle.Text + "," + backcolor.Text + "," + fonttype.Text + "," + cellborder.Text;
            return str;
        }
    }
}
