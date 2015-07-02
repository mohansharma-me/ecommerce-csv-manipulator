using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSVManip
{
    public partial class frmMain : Form
    {
        private String testHTMLFile = "";
        private String browserSettingFile = "browser.ini";
        private String firefoxPath = null, iePath = null, chromePath = null;
        private String _filename = null;
        private int filterIndex = 0;
        private System.Collections.Hashtable rows = new System.Collections.Hashtable();
        private long _lineNumber = 1;

        public frmMain()
        {
            InitializeComponent();
            testHTMLFile = Application.LocalUserAppDataPath + "\\preview.html";
            newFile();
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            
            #region Validate Excel Installation
            frmWait wait = new frmWait("Finding Microsoft Excel...", "Looking for Microsoft Excel...");
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                bool excelFlag = false;
                try
                {
                    Excel.Application app = new Excel.Application();
                    Excel.Workbook workBook = app.Workbooks.Add();
                    workBook.Close(false);
                    app.Quit();
                    excelFlag = true;
                }
                catch (Exception) { excelFlag = false; }

                if (!excelFlag)
                {
                    Invoke(new Action(() => {
                        MessageBox.Show(this, "Sorry, you didn't installed (or installed version is not supported) Microsoft Office.", "Missing Requirement", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                    }));
                    return;
                }

                Invoke(new Action(() => {
                    wait.Close();
                }));

            });
            thread.Priority = System.Threading.ThreadPriority.Highest;
            thread.Start();
            wait.ShowDialog(this);
            #endregion

            #region GetBrowserSettings
            new System.Threading.Thread(() =>
            {
                try
                {
                    if (System.IO.File.Exists(browserSettingFile))
                    {
                        String[] lines = System.IO.File.ReadAllLines(browserSettingFile);
                        foreach (String line in lines)
                        {
                            if (line.StartsWith("firefox"))
                                firefoxPath = line.Substring(7);
                            else if (line.StartsWith("ie"))
                                iePath = line.Substring(2);
                            else if (line.StartsWith("chrome"))
                                chromePath = line.Substring(6);

                            if (line.StartsWith("firefox"))
                            {
                                String exePath=line.Substring(7);
                                Invoke(new Action(() => {
                                    pbPreviewFirefox.Image = pbGoFirefox.Image = null;
                                    pbPreviewFirefox.Image = pbGoFirefox.Image = Icon.ExtractAssociatedIcon(exePath).ToBitmap();
                                }));
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    Invoke(new Action(() => {
                        MessageBox.Show(this, "Error while loading browser settings, please try again." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                } 

            }).Start();
            #endregion

            txtPic1.Focus();
            chkToolbox.Checked = true;
        }

        private void newFile()
        {
            filename = null;
            rows = new System.Collections.Hashtable();
            lineNumber = 1;

            txtPic1.Text = txtPic2.Text = txtPic3.Text = txtPic4.Text = txtPic5.Text = txtPic6.Text = txtPic7.Text = txtPic8.Text = txtPic9.Text = txtPic10.Text = txtPic11.Text = txtPic12.Text = "";
            txtTitle.Text = txtPrice.Text = txtRefNo.Text = txtRefUrl.Text = "";
            txtText.Text = "";
            htmlEditorControl.InnerHtml = "";
            htmlViewer.Text = "";

            btnRefreshPictures_Click(btnRefreshPictures, new EventArgs());
            
            txtPic1.Focus();

            btnTextView_Click(btnTextView, new EventArgs());
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnAddPic(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            int tagNo = int.Parse(ctrl.Tag.ToString());
            TextBox textBox = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Images Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tif;|All files|*.*";
            od.CheckFileExists = true;
            if (od.ShowDialog(this) == DialogResult.OK)
            {
                switch (tagNo)
                {
                    case 1:
                        txtPic1.Text = od.FileName;
                        textBox = txtPic1;
                        break;
                    case 2:
                        txtPic2.Text = od.FileName;
                        textBox = txtPic2;
                        break;
                    case 3:
                        txtPic3.Text = od.FileName;
                        textBox = txtPic3;
                        break;
                    case 4:
                        txtPic4.Text = od.FileName;
                        textBox = txtPic4;
                        break;
                    case 5:
                        txtPic5.Text = od.FileName;
                        textBox = txtPic5;
                        break;
                    case 6:
                        txtPic6.Text = od.FileName;
                        textBox = txtPic6;
                        break;
                    case 7:
                        txtPic7.Text = od.FileName;
                        textBox = txtPic7;
                        break;
                    case 8:
                        txtPic8.Text = od.FileName;
                        textBox = txtPic8;
                        break;
                    case 9:
                        txtPic9.Text = od.FileName;
                        textBox = txtPic9;
                        break;
                    case 10:
                        txtPic10.Text = od.FileName;
                        textBox = txtPic10;
                        break;
                }

                if (textBox != null)
                {
                    txtPic_Changed(textBox, new EventArgs());
                }
            }

        }

        private void btnDelPic(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            int tagNo = int.Parse(ctrl.Tag.ToString());

            PictureBox pb = null;
            TextBox textBox = null;

            switch (tagNo)
            {
                case 1:
                    pb = pictureBox1;
                    txtPic1.Text = "";
                    textBox = txtPic1;
                    break;
                case 2:
                    pb = pictureBox2;
                    txtPic2.Text = "";
                    textBox = txtPic2;
                    break;
                case 3:
                    pb = pictureBox3;
                    txtPic3.Text = "";
                    textBox = txtPic3;
                    break;
                case 4:
                    pb = pictureBox4;
                    txtPic4.Text = "";
                    textBox = txtPic4;
                    break;
                case 5:
                    pb = pictureBox5;
                    txtPic5.Text = "";
                    textBox = txtPic5;
                    break;
                case 6:
                    pb = pictureBox6;
                    txtPic6.Text = "";
                    textBox = txtPic6;
                    break;
                case 7:
                    pb = pictureBox7;
                    txtPic7.Text = "";
                    textBox = txtPic7;
                    break;
                case 8:
                    pb = pictureBox8;
                    txtPic8.Text = "";
                    textBox = txtPic8;
                    break;
                case 9:
                    pb = pictureBox9;
                    txtPic9.Text = "";
                    textBox = txtPic9;
                    break;
            }

            if (textBox != null && pb!=null)
            {
                if (pb.Image != null)
                    pb.Image.Dispose();
                pb.Image = null;
                pb.ImageLocation = null;
                pb.CancelAsync();
                pb.Invalidate();
            }
        }

        private void txtPic_Changed(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            int tagNo = int.Parse(ctrl.Tag.ToString());

            String fileUri = null;
            PictureBox pb = null;

            switch (tagNo)
            {
                case 1:
                    fileUri = txtPic1.Text;
                    pb = pictureBox1;
                    break;

                case 2:
                    fileUri = txtPic2.Text;
                    pb = pictureBox2;
                    break;

                case 3:
                    fileUri = txtPic3.Text;
                    pb = pictureBox3;
                    break;

                case 4:
                    fileUri = txtPic4.Text;
                    pb = pictureBox4;
                    break;

                case 5:
                    fileUri = txtPic5.Text;
                    pb = pictureBox5;
                    break;

                case 6:
                    fileUri = txtPic6.Text;
                    pb = pictureBox6;
                    break;

                case 7:
                    fileUri = txtPic7.Text;
                    pb = pictureBox7;
                    break;

                case 8:
                    fileUri = txtPic8.Text;
                    pb = pictureBox8;
                    break;

                case 9:
                    fileUri = txtPic9.Text;
                    pb = pictureBox9;
                    break;

            }

            if (ctrl.Text.Trim().Length == 0)
            {

                try
                {
                    if (pb.Image != null)
                        pb.Image.Dispose();
                    pb.Image = null;
                    pb.ImageLocation = null;
                    pb.CancelAsync();
                    pb.Invalidate();

                }
                catch (Exception) { }
                return;
            }

            if (fileUri != null && pb != null)
            {
                try
                {
                    if (pb.Image != null)
                        pb.Image.Dispose();
                    pb.Image = null;
                    pb.ImageLocation = null;
                    pb.CancelAsync();
                    pb.Invalidate();
                    
                    System.Threading.Thread thread = new System.Threading.Thread(() => {
                        PictureBox pbTmp = new PictureBox();
                        long curLineNumber = lineNumber;
                        bool success = false;
                        try
                        {
                            pbTmp.Load(fileUri);
                            success = true;
                        }
                        catch (Exception ex)
                        {
                            Invoke(new Action(() =>
                            {
                                MessageBox.Show(this, "Error while loading picture @" + fileUri + ", please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }));
                            success = false;
                        }
                        if (success)
                            Invoke(new Action(() =>
                            {
                                pb.Image = pbTmp.Image;
                            }));
                        else
                            Invoke(new Action(() =>
                            {
                                pb.Image = null;
                            }));
                    });
                    thread.Priority = System.Threading.ThreadPriority.Highest;
                    thread.Start();
                    //pb.LoadAsync(fileUri);
                }
                catch (Exception ex)
                {
                    try
                    {
                        pb.CancelAsync();
                    }
                    catch (Exception) { }
                    MessageBox.Show(this, "Error while loading picture file from uri : '" + fileUri + "' with error message:" + Environment.NewLine + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnRefreshPictures_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null) pictureBox1.Image.Dispose();
            pictureBox1.Image = null;

            txtPic_Changed(txtPic1, new EventArgs());
            txtPic_Changed(txtPic2, new EventArgs());
            txtPic_Changed(txtPic3, new EventArgs());
            txtPic_Changed(txtPic4, new EventArgs());
            txtPic_Changed(txtPic5, new EventArgs());
            txtPic_Changed(txtPic6, new EventArgs());
            txtPic_Changed(txtPic7, new EventArgs());
            txtPic_Changed(txtPic8, new EventArgs());
            txtPic_Changed(txtPic9, new EventArgs());
            txtPic_Changed(txtPic10, new EventArgs());
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            
        }

        /////////////////////////////////////////////////////////////////


        public long lineNumber
        {
            get
            {
                return _lineNumber;
            }

            set
            {
                if (!rows.ContainsKey(value))
                {
                    rows.Add(value, new Row());
                }

                refreshControls(value);

                _lineNumber = value;
                lblCurrentLine.Text = lblCurrentLine1.Text = "CSV: " + (_lineNumber + 1);
            }
        }

        public String filename
        {
            get { return _filename; }
            set
            {
                if (value == null)
                {
                    Text = "CSV-HTML ECOMMERCE LISTER [NewFile]";
                }
                else
                {
                    Text = "CSV-HTML ECOMMERCE LISTER [" + value + "]";
                }
                _filename = value;
            }
        }

        private void refreshControls(long line)
        {
            if (rows.ContainsKey(line))
            {
                Row row=(rows[line]) as Row;
                txtPic1.Text = row.pic1;
                txtPic2.Text = row.pic2;
                txtPic3.Text = row.pic3;
                txtPic4.Text = row.pic4;
                txtPic5.Text = row.pic5;
                txtPic6.Text = row.pic6;
                txtPic7.Text = row.pic7;
                txtPic8.Text = row.pic8;
                txtPic9.Text = row.pic9;
                txtPic10.Text = row.pic10;
                txtPic11.Text = row.pic11;
                txtPic12.Text = row.pic12;

                txtTitle.Text = row.title;
                txtPrice.Text = row.price.ToString();
                HTMLCode=row.text;
                txtRefNo.Text = row.ref_no;
                txtRefUrl.Text = row.ref_url;
                btnRefreshPictures_Click(btnRefreshPictures, new EventArgs());
            }
        }

        private void saveControls(long line)
        {
            if (rows.ContainsKey(line))
            {
                Row row = (rows[line]) as Row;
                row.pic1 = txtPic1.Text;
                row.pic2 = txtPic2.Text;
                row.pic3 = txtPic3.Text;
                row.pic4 = txtPic4.Text;
                row.pic5 = txtPic5.Text;
                row.pic6 = txtPic6.Text;
                row.pic7 = txtPic7.Text;
                row.pic8 = txtPic8.Text;
                row.pic9 = txtPic9.Text;
                row.pic10 = txtPic10.Text;
                row.pic11 = txtPic11.Text;
                row.pic12 = txtPic12.Text;
                double price=0;
                double.TryParse(txtPrice.Text.Trim(), out price);
                row.title = txtTitle.Text;
                row.price = price;
                row.text = HTMLCode;
                //row.text = txtText.Text;
                row.ref_no = txtRefNo.Text;
                row.ref_url = txtRefUrl.Text;
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            newFile();
        }

        private void btnNextRecord_Click(object sender, EventArgs e)
        {
            saveControls(lineNumber);
            lineNumber = lineNumber + 1;
        }

        private void btnPrevRecord_Click(object sender, EventArgs e)
        {
            if (lineNumber > 1)
            {
                saveControls(lineNumber);
                lineNumber = lineNumber - 1;
            }
                
        }

        private void btn1stRC_Click(object sender, EventArgs e)
        {
            lineNumber = 1;
        }

        private void btnLastRC_Click(object sender, EventArgs e)
        {
            lineNumber = rows.Count;
        }

        private void btnSaveCurrent_Click(object sender, EventArgs e)
        {
            saveControls(lineNumber);
            btnSaveCurrent.Text = "Saved";
            new System.Threading.Thread(() => {
                try
                {
                    System.Threading.Thread.Sleep(1000);
                    Invoke(new Action(() => {
                        btnSaveCurrent.Text = "&Save current";
                    }));
                }
                catch (Exception) { }
            }).Start();
        }

        private void btnLineGo_Click(object sender, EventArgs e)
        {
            long usrLine = 0;

            if (long.TryParse(txtCSVLineNo.Text, out usrLine))
            {
                if (usrLine > 1)
                {
                    lineNumber = usrLine - 1;
                }
            }
        }

        private void btnTextView_Click(object sender, EventArgs e)
        {
            htmlEditorControl.Visible = htmlViewer.Visible = false;
            txtText.Visible = true;
            htmlEditorControl.Dock = htmlViewer.Dock = DockStyle.None;
            txtText.Dock = DockStyle.Fill;
            txtText.BringToFront();
            if (htmlEditorControl.InnerHtml != null)
                txtText.Text = htmlEditorControl.InnerHtml;
        }

        private void btnHtmlView_Click(object sender, EventArgs e)
        {
            txtText.Visible = htmlViewer.Visible = false;
            htmlEditorControl.Visible = true;
            txtText.Dock = htmlViewer.Dock = DockStyle.None;
            htmlEditorControl.Dock = DockStyle.Fill;
            htmlEditorControl.BringToFront();
            if (txtText.Text.Trim().Length > 0)
                htmlEditorControl.InnerHtml = txtText.Text;
        }

        private void btnPreviewView_Click(object sender, EventArgs e)
        {
            bool isTextVisible=txtText.Visible;
            bool isHtmlEditorVisible=htmlEditorControl.Visible;

            htmlEditorControl.Visible = !isTextVisible;
            htmlEditorControl.Dock = isHtmlEditorVisible ? DockStyle.Left : DockStyle.None;

            txtText.Visible = !isHtmlEditorVisible;
            txtText.Dock = txtText.Visible ? DockStyle.Left : DockStyle.None;

            htmlViewer.Visible = true;
            htmlViewer.Dock = DockStyle.Right;
        }

        private void txtText_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtText_Leave(object sender, EventArgs e)
        {
            try
            {
                //htmlViewer.Text = txtText.Text;
                //htmlViewer.PerformLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while refreshing html-rendering layer, please try again." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRefURLGo_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRefUrl.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while navigating to your provided reference link." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public String ExtFilter
        {
            get
            {
                return "Excel Workbook|*.xlsx|Excel Macro-enabled Workbook|*.xlsm|Excel 97-2003 Workbook|*.xls|Excel Macro-enabled Template|*.xltm|Excel 97-2003 Template|*.xlt|Text (Tab Limited)|*.txt|CSV (Comma separated)|*.csv|CSV (Macintosh)|*.csv|CSV (MS-Dos)|*.csv|All files|*.*";
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename != null)
            {
                if (MessageBox.Show(this, "Are you want to save currently opened file ?", "Save prompt", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    saveAsToolStripMenuItem_Click(sender, e);
                    newFile();
                }
            }

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = ExtFilter;
            if (od.ShowDialog(this) == DialogResult.OK)
            {
                rows = new System.Collections.Hashtable();
                _lineNumber = 0;
                filename = od.FileName;
                filterIndex = od.FilterIndex;
                bool success = false;

                frmWait wait = new frmWait("Opening document...", "Please wait...");
                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                    try
                    {
                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook wb = excelApp.Workbooks.Open(filename);

                        foreach (Excel.Worksheet sheet in wb.Worksheets)
                        {
                            
                            Excel.Range usedRange = sheet.UsedRange;

                            if (usedRange.Columns.Count >= 17)
                            {
                                for (long row = 2; row <= usedRange.Rows.Count; row++)
                                {
                                    try
                                    {
                                        Row tmpRow = new Row();
                                        tmpRow.pic1 = usedRange.get_Range("A" + row).get_Value() == null ? "" : usedRange.get_Range("A" + row).get_Value().ToString();
                                        tmpRow.pic2 = usedRange.get_Range("B" + row).get_Value() == null ? "" : usedRange.get_Range("B" + row).get_Value().ToString();
                                        tmpRow.pic3 = usedRange.get_Range("C" + row).get_Value() == null ? "" : usedRange.get_Range("C" + row).get_Value().ToString();
                                        tmpRow.pic4 = usedRange.get_Range("D" + row).get_Value() == null ? "" : usedRange.get_Range("D" + row).get_Value().ToString();
                                        tmpRow.pic5 = usedRange.get_Range("E" + row).get_Value() == null ? "" : usedRange.get_Range("E" + row).get_Value().ToString();
                                        tmpRow.pic6 = usedRange.get_Range("F" + row).get_Value() == null ? "" : usedRange.get_Range("F" + row).get_Value().ToString();
                                        tmpRow.pic7 = usedRange.get_Range("G" + row).get_Value() == null ? "" : usedRange.get_Range("G" + row).get_Value().ToString();
                                        tmpRow.pic8 = usedRange.get_Range("H" + row).get_Value() == null ? "" : usedRange.get_Range("H" + row).get_Value().ToString();
                                        tmpRow.pic9 = usedRange.get_Range("I" + row).get_Value() == null ? "" : usedRange.get_Range("I" + row).get_Value().ToString();
                                        tmpRow.pic10 = usedRange.get_Range("J" + row).get_Value() == null ? "" : usedRange.get_Range("J" + row).get_Value().ToString();

                                        tmpRow.pic11 = usedRange.get_Range("K" + row).get_Value() == null ? "" : usedRange.get_Range("K" + row).get_Value().ToString();
                                        tmpRow.pic12 = usedRange.get_Range("L" + row).get_Value() == null ? "" : usedRange.get_Range("L" + row).get_Value().ToString();

                                        tmpRow.title = usedRange.get_Range("M" + row).get_Value() == null ? "" : usedRange.get_Range("M" + row).get_Value().ToString();
                                        tmpRow.price = usedRange.get_Range("N" + row).get_Value() == null ? 0 : usedRange.get_Range("N" + row).get_Value();
                                        tmpRow.text = usedRange.get_Range("O" + row).get_Value() == null ? "" : usedRange.get_Range("O" + row).get_Value().ToString();
                                        tmpRow.ref_no = usedRange.get_Range("P" + row).get_Value() == null ? "" : usedRange.get_Range("P" + row).get_Value().ToString();
                                        tmpRow.ref_url = usedRange.get_Range("Q" + row).get_Value() == null ? "" : usedRange.get_Range("Q" + row).get_Value().ToString();

                                        if (!tmpRow.isEmpty)
                                        {
                                            rows.Add(row - 1, tmpRow);
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Invoke(new Action(() =>
                                        {
                                            MessageBox.Show(this, "Error while parsing data from opened document, skipping current row and heading to next."+Environment.NewLine+"Your sheet must have data for price column as numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        }));
                                    }

                                }
                            }
                            else
                            {
                                if (usedRange.Rows.Count == 1 && usedRange.Columns.Count == 1)
                                {

                                }
                                else
                                {
                                    Invoke(new Action(() =>
                                    {
                                        MessageBox.Show(this, "Sheet:[" + sheet.Name + "] doesn't contain valid data, moving to next sheet."+Environment.NewLine+"Your sheet must contain at least 17 columns in order as shown in given template file.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }));
                                }
                            }

                        }

                        killExcel(wb.Worksheets, wb, excelApp.Workbooks, excelApp);
                        Invoke(new Action(() => {
                            lineNumber = 1;
                            wait.Close();
                        }));

                    }
                    catch (Exception ex)
                    {
                        Invoke(new Action(() => {
                            MessageBox.Show(this, "Error while opening document, please try again." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                });
                thread.Priority = System.Threading.ThreadPriority.Highest;
                thread.Start();
                wait.ShowDialog(this);


            }
        }

        private bool saveAs = false;

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename == null || saveAs)
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Filter = ExtFilter;
                sd.OverwritePrompt = true;
                if (sd.ShowDialog(this) == DialogResult.OK)
                {
                    filename = sd.FileName;
                    filterIndex = sd.FilterIndex;
                }
                else
                {
                    return;
                }
            }

            saveAs = false;

            frmWait wait = new frmWait("Saving document...", "Writing, please wait...");
            wait.pb.Maximum = rows.Count;
            wait.pb.Value = 0;
            wait.pb.Style = ProgressBarStyle.Continuous;

            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                try
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook wb = excelApp.Workbooks.Add();
                    Excel.Worksheet sheet = wb.Worksheets.Add();
                    sheet.Name = "DataSheet";

                    sheet.get_Range("A1").set_Value(Type.Missing, "PIC1");
                    sheet.get_Range("B1").set_Value(Type.Missing, "PIC2");
                    sheet.get_Range("C1").set_Value(Type.Missing, "PIC3");
                    sheet.get_Range("D1").set_Value(Type.Missing, "PIC4");
                    sheet.get_Range("E1").set_Value(Type.Missing, "PIC5");
                    sheet.get_Range("F1").set_Value(Type.Missing, "PIC6");
                    sheet.get_Range("G1").set_Value(Type.Missing, "PIC7");
                    sheet.get_Range("H1").set_Value(Type.Missing, "PIC8");
                    sheet.get_Range("I1").set_Value(Type.Missing, "PIC9");
                    sheet.get_Range("J1").set_Value(Type.Missing, "PIC10");

                    sheet.get_Range("K1").set_Value(Type.Missing, "PIC11");
                    sheet.get_Range("L1").set_Value(Type.Missing, "PIC12");

                    sheet.get_Range("M1").set_Value(Type.Missing, "TITLE");
                    sheet.get_Range("N1").set_Value(Type.Missing, "PRICE");
                    sheet.get_Range("O1").set_Value(Type.Missing, "DESCRIPTION");
                    sheet.get_Range("P1").set_Value(Type.Missing, "REF NO");
                    sheet.get_Range("Q1").set_Value(Type.Missing, "REF URL");

                    foreach(long _line in rows.Keys) {
                        Row row = (Row)rows[_line];
                        if (row == null) continue;
                        if (row != null && row.isEmpty) continue;
                        long line = _line + 1;
                        sheet.get_Range("A" + line).set_Value(Type.Missing, row.pic1);
                        sheet.get_Range("B" + line).set_Value(Type.Missing, row.pic2);
                        sheet.get_Range("C" + line).set_Value(Type.Missing, row.pic3);
                        sheet.get_Range("D" + line).set_Value(Type.Missing, row.pic4);
                        sheet.get_Range("E" + line).set_Value(Type.Missing, row.pic5);
                        sheet.get_Range("F" + line).set_Value(Type.Missing, row.pic6);
                        sheet.get_Range("G" + line).set_Value(Type.Missing, row.pic7);
                        sheet.get_Range("H" + line).set_Value(Type.Missing, row.pic8);
                        sheet.get_Range("I" + line).set_Value(Type.Missing, row.pic9);
                        sheet.get_Range("J" + line).set_Value(Type.Missing, row.pic10);

                        sheet.get_Range("K" + line).set_Value(Type.Missing, row.pic11);
                        sheet.get_Range("L" + line).set_Value(Type.Missing, row.pic12);

                        sheet.get_Range("M" + line).set_Value(Type.Missing, row.title);
                        sheet.get_Range("N" + line).set_Value(Type.Missing, row.price);
                        sheet.get_Range("O" + line).set_Value(Type.Missing, row.text);
                        sheet.get_Range("P" + line).set_Value(Type.Missing, row.ref_no);
                        sheet.get_Range("Q" + line).set_Value(Type.Missing, row.ref_url);

                        Invoke(new Action(() =>
                        {
                            wait.pb.Value++;
                        }));
                    }

                    object delimiter=Type.Missing;
                    Excel.XlFileFormat fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook;

                    switch (filterIndex-1)
                    {
                        case 0:
                            fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook;
                            break;
                        case 1:
                            fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                            break;
                        case 2:
                            fileFormat = Excel.XlFileFormat.xlExcel9795;
                            break;
                        case 3:
                            fileFormat = Excel.XlFileFormat.xlOpenXMLTemplateMacroEnabled;
                            break;
                        case 4:
                            fileFormat = Excel.XlFileFormat.xlOpenXMLTemplate;
                            break;
                        case 5:
                            fileFormat = Excel.XlFileFormat.xlTextWindows;
                            delimiter = "\t";
                            break;
                        case 6:
                            fileFormat = Excel.XlFileFormat.xlCSV;
                            delimiter = ",";
                            break;
                        case 7:
                            fileFormat = Excel.XlFileFormat.xlCSVMac;
                            break;
                        case 8:
                            fileFormat = Excel.XlFileFormat.xlCSVMSDOS;
                            break;
                    }

                    wb.SaveAs(filename, fileFormat);
                    wb.Close(true);
                    //excelApp.Quit();

                    try
                    {
                        killExcel(wb.Worksheets, wb, excelApp.Workbooks, excelApp);
                    }
                    catch (Exception) { }
                    

                    Invoke(new Action(() => {
                        wait.Close();
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show(this, "Error while writing to file, please try again." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
            thread.Priority = System.Threading.ThreadPriority.Highest;
            thread.Start();

            wait.ShowDialog(this);

        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveAs = true;
            saveToolStripMenuItem_Click(sender, e);
        }

        private void killExcel(Excel.Sheets _workSheets, Excel.Workbook _workBook, Excel.Workbooks _workBooks, Excel.Application _excelApp)
        {
            try
            {
                try
                {
                    _workBook.Close();
                }
                catch (Exception) { }


                try
                {
                    _workBooks.Close();
                }
                catch (Exception) { }

                try
                {
                    _excelApp.Quit();
                }
                catch (Exception) { }
                
                
            }
            catch (Exception) { }

            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.FinalReleaseComObject(_workSheets);
                Marshal.FinalReleaseComObject(_workBook);
                Marshal.FinalReleaseComObject(_workBooks);
                Marshal.FinalReleaseComObject(_excelApp);
            }
            catch (Exception) { }
        }

        private void txtTitle_TextChanged(object sender, EventArgs e)
        {
            lblTitleCount.Text = "[" + txtTitle.Text.Length + "]";
        }

        private void pbGoFirefox_Click(object sender, EventArgs e)
        {
            if (firefoxPath == null)
            {
                firefoxToolStripMenuItem_Click(sender, e);
                if (firefoxPath == null) return;
            }
            
            try
            {
                System.Diagnostics.Process.Start(firefoxPath,txtRefUrl.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while navigating to your provided reference link." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void firefoxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            storeBrowserSettings("firefox");
        }

        private void storeBrowserSettings(String saveName)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Executable file|*.exe|All files|*.*";
            if (open.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    switch (saveName)
                    {
                        case "firefox":
                            firefoxPath = open.FileName;
                            break;
                        case "chrome":
                            chromePath = open.FileName;
                            break;
                        case "ie":
                            iePath = open.FileName;
                            break;
                    }
                    System.IO.File.AppendAllLines(browserSettingFile, new List<String>() { saveName + open.FileName });
                    String exePath = open.FileName;
                    Invoke(new Action(() =>
                    {
                        pbPreviewFirefox.Image = pbGoFirefox.Image = null;
                        pbPreviewFirefox.Image = pbGoFirefox.Image = Icon.ExtractAssociatedIcon(exePath).ToBitmap();
                    }));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error while writing settings to file, please try again." + Environment.NewLine + "Error message:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void internetExplorerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            storeBrowserSettings("ie");
        }

        private void chromeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            storeBrowserSettings("chrome");
        }

        private void pbGoIE_Click(object sender, EventArgs e)
        {
            if (iePath == null)
            {
                internetExplorerToolStripMenuItem_Click(sender, e);
                if (iePath == null) return;
            }

            try
            {
                System.Diagnostics.Process.Start(iePath, txtRefUrl.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while navigating to your provided reference link." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pbGoChrome_Click(object sender, EventArgs e)
        {
            if (chromePath == null)
            {
                chromeToolStripMenuItem_Click(sender, e);
                if (chromePath == null) return;
            }

            try
            {
                System.Diagnostics.Process.Start(chromePath, txtRefUrl.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while navigating to your provided reference link." + Environment.NewLine + "Error msg:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public String HTMLCode
        {
            get 
            {
                if (htmlEditorControl.Visible)
                    return htmlEditorControl.InnerHtml == null ? "" : htmlEditorControl.InnerHtml;
                else if (txtText.Visible)
                    return txtText.Text;
                else
                    return "n/a";
            }
            
            set 
            {
                htmlViewer.Text = value;
                htmlEditorControl.InnerHtml = value;
                txtText.Text = value;
            }
        }

        private void pbPreviewFirefox_Click(object sender, EventArgs e)
        {
            try
            {
                if (firefoxPath == null)
                {
                    firefoxToolStripMenuItem_Click(sender, e);
                    if (firefoxPath == null) return;
                }

                if (HTMLCode==null)
                {
                    MessageBox.Show(this, "Sorry, there is no content in html editor.", "No Content", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    System.IO.File.WriteAllText(testHTMLFile, HTMLCode);
                    System.Diagnostics.Process.Start(firefoxPath, "\"" + testHTMLFile + "\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while storing html data temporarily and showing it on browser, please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pbPreviewIE_Click(object sender, EventArgs e)
        {
            try
            {
                if (iePath == null)
                {
                    internetExplorerToolStripMenuItem_Click(sender, e);
                    if (iePath == null) return;
                }

                if (htmlEditorControl.InnerHtml == null)
                {
                    MessageBox.Show(this, "Sorry, there is no content in html editor.", "No Content", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    System.IO.File.WriteAllText("temp.html", htmlEditorControl.InnerHtml);
                    System.Diagnostics.Process.Start(iePath, Application.StartupPath + "\\temp.html");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while storing html data temporarily and showing it on browser, please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pbPreviewChrome_Click(object sender, EventArgs e)
        {
            try
            {
                if (chromePath == null)
                {
                    chromeToolStripMenuItem_Click(sender, e);
                    if (chromePath == null) return;
                }

                if (htmlEditorControl.InnerHtml == null)
                {
                    MessageBox.Show(this, "Sorry, there is no content in html editor.", "No Content", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    System.IO.File.WriteAllText("temp.html", htmlEditorControl.InnerHtml);
                    System.Diagnostics.Process.Start(chromePath, Application.StartupPath + "\\temp.html");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while storing html data temporarily and showing it on browser, please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkToolbox_CheckedChanged(object sender, EventArgs e)
        {
            htmlEditorControl.ToolbarVisible = chkToolbox.Checked;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            new frmImgViewer((sender as PictureBox).Image).ShowDialog(this);
        }

        private void frmMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtCSVLineNo.Focused && e.KeyCode == Keys.Enter)
            {
                btnLineGo_Click(btnLineGo, new EventArgs());
            }
            else if (txtRefUrl.Focused && e.KeyCode == Keys.Enter)
            {
                btnGoButton_Click(btnGoButton, new EventArgs());
            }
            else if (e.Shift && e.KeyCode == Keys.S)
            {
                btnSaveCurrent_Click(btnSaveCurrent, new EventArgs());
            }
            else
            {
                if (panel1.Focused && e.KeyCode == Keys.Right)
                {
                    btnNextRecord_Click(btnNextRecord, new EventArgs());
                }
                else if (panel1.Focused && e.KeyCode == Keys.Left)
                {
                    btnPrevRecord_Click(btnPrevRecord, new EventArgs());
                }
                else if (panel1.Focused && e.KeyCode == Keys.Up)
                {
                    try
                    {
                        panel1.VerticalScroll.Value -= 15;
                    }
                    catch (Exception) { panel1.VerticalScroll.Value = 0; }
                }
                else if (panel1.Focused && e.KeyCode == Keys.Down)
                {
                    try
                    {
                        panel1.VerticalScroll.Value += 15;
                    }
                    catch (Exception) { panel1.VerticalScroll.Value = panel1.VerticalScroll.Maximum; }
                }
            }
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            bool isTextVisible = txtText.Visible;
            bool isHtmlEditorVisible = htmlEditorControl.Visible;
            
            htmlEditorControl.Visible = !isTextVisible;
            htmlEditorControl.Dock = isHtmlEditorVisible ? DockStyle.Left : DockStyle.None;
            htmlEditorControl.Width = htmlEditorControl.Visible ? panel2.Width / 2 : htmlEditorControl.Width;

            txtText.Visible = !isHtmlEditorVisible;
            txtText.Dock = txtText.Visible ? DockStyle.Left : DockStyle.None;
            txtText.Width = txtText.Visible ? panel2.Width / 2 : txtText.Width;

            if (htmlEditorControl.InnerHtml != null)
                txtText.Text = htmlEditorControl.InnerHtml;

            htmlViewer.Visible = true;
            htmlViewer.Dock = DockStyle.Right;
            htmlViewer.Text = txtText.Text;


            htmlViewer.Width = (panel2.Width / 2) - 10;

            if (chkFullPreview.Checked)
            {
                htmlEditorControl.Visible = false;
                txtText.Visible = false;
                htmlViewer.Visible = true;
                htmlViewer.Dock = DockStyle.Fill;
            }
        }

        private void htmlEditorControl_Leave(object sender, EventArgs e)
        {
            if (htmlEditorControl.InnerHtml != null)
            {
                htmlViewer.Text = htmlEditorControl.InnerHtml;
                txtText.Text = htmlEditorControl.InnerHtml;
            }
        }

        private void htmlViewer_Click(object sender, EventArgs e)
        {
            htmlEditorControl_Leave(sender, e);
        }

        private void btnGoButton_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(txtRefUrl.Text);
        }

        private void browserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            storeBrowserSettings("firefox");
        }

        private void txtText_Leave_1(object sender, EventArgs e)
        {
            htmlViewer.Text = txtText.Text;
            htmlEditorControl.InnerHtml = txtText.Text;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void panel1_Click(object sender, EventArgs e)
        {
            panel1.Focus();
        }

        private void saveBasicCSVTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "CSV Template File";
            sfd.Filter = "CSV File|*.csv|All Files|*.*";
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    System.IO.File.WriteAllText(sfd.FileName, "PIC1,PIC2,PIC3,PIC4,PIC5,PIC6,PIC7,PIC8,PIC9,PIC10,PIC11,PIC12,TITLE,PRICE,DESCRIPTION,REF NO,REF URL1");
                    MessageBox.Show(this, "Template file saved successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error while writing data to given file, please try again." + Environment.NewLine + "Error message:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnSearchREFNo_Click(object sender, EventArgs e)
        {
            String searchKey = txtRefNo.Text.Trim().ToLower();
            frmWait wait = new frmWait("Searching...", "Please wait while searching entry by refference number...");
            System.Threading.Thread thread = new System.Threading.Thread(() => {

                try
                {
                    foreach (long key in rows.Keys)
                    {
                        Row row = (Row)rows[key];
                        if (row != null)
                        {
                            if (row.ref_no.Trim().ToLower().Equals(searchKey))
                            {
                                lineNumber = key;
                                break;
                            }
                        }
                    }
                }
                catch (Exception) { }

                try
                {
                    Invoke(new Action(() =>
                    {
                        wait.Close();
                    }));
                }
                catch (Exception) { }

            });
            thread.Priority = System.Threading.ThreadPriority.Highest;
            thread.Start();
            wait.ShowDialog(this);
        }

        private void txtCSVLineNo_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtText_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu contextMenu = new ContextMenu();
                contextMenu.MenuItems.Add(new MenuItem("&Cut", __cut));
                contextMenu.MenuItems.Add(new MenuItem("&Copy", __copy));
                contextMenu.MenuItems.Add(new MenuItem("&Paste", __paste));
                contextMenu.MenuItems.Add(new MenuItem("&Select All", __selectAll));
                contextMenu.MenuItems.Add(new MenuItem("-"));
                contextMenu.MenuItems.Add(new MenuItem("&Font", __font));
                contextMenu.MenuItems.Add(new MenuItem("Find && &Replace", __findReplace));

                contextMenu.MenuItems[0].Shortcut = Shortcut.CtrlX;
                contextMenu.MenuItems[1].Shortcut = Shortcut.CtrlC;
                contextMenu.MenuItems[2].Shortcut = Shortcut.CtrlV;
                contextMenu.MenuItems[3].Shortcut = Shortcut.CtrlA;
                
                contextMenu.MenuItems[0].Shortcut = Shortcut.CtrlF;
                contextMenu.MenuItems[1].Shortcut = Shortcut.CtrlH;

                contextMenu.Show(sender as Control, e.Location);
            }
        }

        private void __cut(object sender, EventArgs e)
        {
            txtText.Cut();
        }

        private void __copy(object sender, EventArgs e)
        {
            txtText.Copy();
        }

        private void __paste(object sender, EventArgs e)
        {
            txtText.Paste();
        }

        private void __selectAll(object sender, EventArgs e)
        {
            txtText.SelectAll();
        }

        private void __font(object sender, EventArgs e)
        {
            FontDialog fontDialog = new FontDialog();
            if (fontDialog.ShowDialog(this) == DialogResult.OK)
            {
                txtText.Font = fontDialog.Font;
            }
        }

        private void __findReplace(object sender, EventArgs e)
        {
            frmFindReplace findReplace = new frmFindReplace();
            if (findReplace.ShowDialog(this) == DialogResult.OK)
            {
                if (findReplace.txtFind.Text.Length > 0)
                {
                    txtText.Text = txtText.Text.Replace(findReplace.txtFind.Text, findReplace.txtReplace.Text);
                }
            }
        }

        private void hiddenMenuToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void txtText_Enter(object sender, EventArgs e)
        {
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            __cut(sender, e);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            __copy(sender, e);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            __paste(sender, e);
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            __selectAll(sender, e);
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            __font(sender, e);
        }

        private void fingReplaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            __findReplace(sender, e);
        }

        private void txtText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.X)
            {
                __cut(sender, new EventArgs());
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                __copy(sender, new EventArgs());
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                //__paste(sender, new EventArgs());
            }
            else if (e.Control && e.KeyCode == Keys.A)
            {
                __selectAll(sender, new EventArgs());
            }
            else if (e.Control && e.KeyCode == Keys.F)
            {
                __font(sender, new EventArgs());
            }
            else if (e.Control && e.KeyCode == Keys.H)
            {
                __findReplace(sender, new EventArgs());
            }
            else
            {
                e.Handled = false;
            }
        }
        
        

    }
}
