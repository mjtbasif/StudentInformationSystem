using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace StudentInfoSystem
{
    public partial class Form1 : Form
    {
        string path = "";
        public Form1()
        {
            InitializeComponent();
            clearLabels();
        }
        void clearLabels() //clearing all label
        {
            name.Text = "";
            roll.Text = "";
            merit.Text = "";
            grade.Text = "";
            marks.Text = "";
            gpa.Text = "";
            //Bangla
            oBan.Text = "";
            sBan.Text = "";
            pracBan.Text = "";
            tBan.Text = "";
            gBan.Text = "";
            gpaBan.Text = "";

            //English 
            oEng.Text = "";
            sEng.Text = "";
            pracEng.Text = "";
            tEng.Text = "";
            gEng.Text = "";
            gpaEng.Text = "";

            //Mathematics
            oMat.Text = "";
            sMat.Text = "";
            pracMat.Text = "";
            tMat.Text = "";
            gMat.Text = "";
            gpaMat.Text = "";

            //Religion
            oRel.Text = "";
            sRel.Text = "";
            pracRel.Text = "";
            tRel.Text = "";
            gRel.Text = "";
            gpaRel.Text = "";

            //Bangladesh and Global Studies			
            oBnG.Text = "";
            sBnG.Text = "";
            pracBnG.Text = "";
            tBnG.Text = "";
            gBnG.Text = "";
            gpaBnG.Text = "";

            //Physical Health, Health Science and Games and Sports
            oPHS.Text = "";
            sPHS.Text = "";
            pracPHS.Text = "";
            tPHS.Text = "";
            gPHS.Text = "";
            gpaPHS.Text = "";

            //Information & Communication Technology
            oICT.Text = "";
            sICT.Text = "";
            pracICT.Text = "";
            tICT.Text = "";
            gICT.Text = "";
            gpaICT.Text = "";

            //Career Education
            oCE.Text = "";
            sCE.Text = "";
            pracCE.Text = "";
            tCE.Text = "";
            gCE.Text = "";
            gpaCE.Text = "";

            //Physics
            oPhy.Text = "";
            sPhy.Text = "";
            pracPhy.Text = "";
            tPhy.Text = "";
            gPhy.Text = "";
            gpaPhy.Text = "";

            //Chemistry
            oChe.Text = "";
            sChe.Text = "";
            pracChe.Text = "";
            tChe.Text = "";
            gChe.Text = "";
            gpaChe.Text = "";

            //Biology
            oBio.Text = "";
            sBio.Text = "";
            pracBio.Text = "";
            tBio.Text = "";
            gBio.Text = "";
            gpaBio.Text = "";

            //Higher Math
            oHM.Text = "";
            sHM.Text = "";
            pracHM.Text = "";
            tHM.Text = "";
            gHM.Text = "";
            gpaHM.Text = "";
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "*.xls|*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    path = ofd.FileName;
                }
            }
                
            if (path == "")
            {
                status.Text = "No file selected";
                filePath.Text = status.Text;
            }
            else
            {
                status.Text = "";
                filePath.Text = path;
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("How to use the software\n\n1. Browse a valid file           \n2. Enter Roll         \n3. Search         \n\nIf the application shows any error, make sure to do a quick repair of your Microsoft Office from control panel\n\nThank you.", "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Developed by\n\n" +
               "Mujtaba Asif\n" +
               "Computer Science and Engineering\n" +
               "East West University\n\n" +
               "Thank you\n", "Hello", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void search_Click_1(object sender, EventArgs e)
        {
            if (path == "")
            {
                status.Text = "no file selected";
                filePath.Text = status.Text;
            }
            else
            {
                status.Text = "searching...";
                Excel.Application xApp = new Excel.Application();
                xApp.Visible = false;
                Excel.Workbook xBook = xApp.Workbooks.Open(path);
                Excel.Worksheet xSheet = xApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                
                Excel.Range userRange = xSheet.UsedRange;
                int i, f = 0;
                for (i = 2; i <= userRange.Rows.Count; i++)
                {
                    if ((xSheet.Cells[i, 2] as Excel.Range).Value != null)
                    {
                        if (skey.Text == (xSheet.Cells[i, 2] as Excel.Range).Value.ToString())
                        {
                            f = 1;
                            break;
                        }
                    }
                }
                if ((xSheet.Cells[i, 2] as Excel.Range).Value != null && f==1)
                {
                    int nf = 1;
                    status.Text = "";
                    //Name
                    if ((xSheet.Cells[i, 3]).Value == null)
                    {

                        name.Text = "";
                        status.Text = "roll found but no information";
                        nf = 0;
                        clearLabels();
                    }
                    else
                    {
                        name.Text = (xSheet.Cells[i, 3] as Excel.Range).Value.ToString();
                    }
                    //Roll
                    if ((xSheet.Cells[i, 2]).Value == null)
                    {

                        roll.Text = "";
                    }
                    else
                    {
                        if (nf == 1)
                            roll.Text = (xSheet.Cells[i, 2] as Excel.Range).Value.ToString();
                        else
                            roll.Text = "";
                    }
                    //Merit
                    if ((xSheet.Cells[i, 1]).Value == null)
                    {

                        merit.Text = "";
                    }
                    else
                    {
                        if (nf == 1)
                            merit.Text = (xSheet.Cells[i, 1] as Excel.Range).Value.ToString();
                        else
                            merit.Text = "";
                    }
                    //Marks
                    if ((xSheet.Cells[i, 1]).Value == null)
                    {

                        marks.Text = "";
                    }
                    else
                    {
                        if (nf == 1)
                            marks.Text = (xSheet.Cells[i, 77] as Excel.Range).Value.ToString();
                        else
                            marks.Text = "";
                    }
                    //Grade
                    if ((xSheet.Cells[i, 1]).Value == null)
                    {

                        grade.Text = "";
                    }
                    else
                    {
                        if (nf == 1)
                            grade.Text = (xSheet.Cells[i, 75] as Excel.Range).Value.ToString();
                        else
                            grade.Text = "";
                    }
                    //GPA
                    if ((xSheet.Cells[i, 1]).Value == null)
                    {

                        gpa.Text = "";
                    }
                    else
                    {
                        if (nf == 1)
                            gpa.Text = (xSheet.Cells[i, 74] as Excel.Range).Value.ToString("0.00");
                        else
                            gpa.Text = "";
                    }

                    if (nf == 1)
                    {
                        //Bangla
                        oBan.Text = (xSheet.Cells[i, 9] as Excel.Range).Value.ToString();
                        sBan.Text = (xSheet.Cells[i, 10] as Excel.Range).Value.ToString();
                        pracBan.Text = "-";
                        tBan.Text = (xSheet.Cells[i, 11] as Excel.Range).Value.ToString();
                        gBan.Text = (xSheet.Cells[i, 12] as Excel.Range).Value.ToString();
                        gpaBan.Text = (xSheet.Cells[i, 13] as Excel.Range).Value.ToString();

                        //English 
                        oEng.Text = "-";
                        sEng.Text = "-";
                        pracEng.Text = "-";
                        tEng.Text = (xSheet.Cells[i, 16] as Excel.Range).Value.ToString();
                        gEng.Text = (xSheet.Cells[i, 17] as Excel.Range).Value.ToString();
                        gpaEng.Text = (xSheet.Cells[i, 18] as Excel.Range).Value.ToString();

                        //Mathematics
                        oMat.Text = (xSheet.Cells[i, 19] as Excel.Range).Value.ToString();
                        sMat.Text = (xSheet.Cells[i, 20] as Excel.Range).Value.ToString();
                        pracMat.Text = "-";
                        tMat.Text = (xSheet.Cells[i, 21] as Excel.Range).Value.ToString();
                        gMat.Text = (xSheet.Cells[i, 22] as Excel.Range).Value.ToString();
                        gpaMat.Text = (xSheet.Cells[i, 23] as Excel.Range).Value.ToString();

                        //Religion
                        oRel.Text = (xSheet.Cells[i, 24] as Excel.Range).Value.ToString();
                        sRel.Text = (xSheet.Cells[i, 25] as Excel.Range).Value.ToString();
                        pracRel.Text = "-";
                        tRel.Text = (xSheet.Cells[i, 26] as Excel.Range).Value.ToString();
                        gRel.Text = (xSheet.Cells[i, 27] as Excel.Range).Value.ToString();
                        gpaRel.Text = (xSheet.Cells[i, 28] as Excel.Range).Value.ToString();

                        //Bangladesh and Global Studies			
                        oBnG.Text = (xSheet.Cells[i, 29] as Excel.Range).Value.ToString();
                        sBnG.Text = (xSheet.Cells[i, 30] as Excel.Range).Value.ToString();
                        tBnG.Text = (xSheet.Cells[i, 31] as Excel.Range).Value.ToString();
                        pracBnG.Text = "-";
                        gBnG.Text = (xSheet.Cells[i, 32] as Excel.Range).Value.ToString();
                        gpaBnG.Text = (xSheet.Cells[i, 33] as Excel.Range).Value.ToString();

                        //Physical Health, Health Science and Games and Sports
                        oPHS.Text = (xSheet.Cells[i, 34] as Excel.Range).Value.ToString();
                        sPHS.Text = (xSheet.Cells[i, 35] as Excel.Range).Value.ToString();
                        pracPHS.Text = (xSheet.Cells[i, 36] as Excel.Range).Value.ToString();
                        tPHS.Text = (xSheet.Cells[i, 37] as Excel.Range).Value.ToString();
                        gPHS.Text = (xSheet.Cells[i, 38] as Excel.Range).Value.ToString();
                        gpaPHS.Text = (xSheet.Cells[i, 39] as Excel.Range).Value.ToString();

                        //Information & Communication Technology
                        oICT.Text = (xSheet.Cells[i, 40] as Excel.Range).Value.ToString();
                        sICT.Text = "-";
                        pracICT.Text = (xSheet.Cells[i, 41] as Excel.Range).Value.ToString();
                        tICT.Text = (xSheet.Cells[i, 42] as Excel.Range).Value.ToString();
                        gICT.Text = (xSheet.Cells[i, 43] as Excel.Range).Value.ToString();
                        gpaICT.Text = (xSheet.Cells[i, 44] as Excel.Range).Value.ToString();

                        //Career Education
                        oCE.Text = (xSheet.Cells[i, 45] as Excel.Range).Value.ToString();
                        sCE.Text = "-";
                        pracCE.Text = (xSheet.Cells[i, 46] as Excel.Range).Value.ToString();
                        tCE.Text = (xSheet.Cells[i, 47] as Excel.Range).Value.ToString();
                        gCE.Text = (xSheet.Cells[i, 48] as Excel.Range).Value.ToString();
                        gpaCE.Text = (xSheet.Cells[i, 49] as Excel.Range).Value.ToString();

                        //Physics
                        oPhy.Text = (xSheet.Cells[i, 50] as Excel.Range).Value.ToString();
                        sPhy.Text = (xSheet.Cells[i, 51] as Excel.Range).Value.ToString();
                        pracPhy.Text = (xSheet.Cells[i, 52] as Excel.Range).Value.ToString();
                        tPhy.Text = (xSheet.Cells[i, 53] as Excel.Range).Value.ToString();
                        gPhy.Text = (xSheet.Cells[i, 54] as Excel.Range).Value.ToString();
                        gpaPhy.Text = (xSheet.Cells[i, 55] as Excel.Range).Value.ToString();

                        //Chemistry
                        oChe.Text = (xSheet.Cells[i, 56] as Excel.Range).Value.ToString();
                        sChe.Text = (xSheet.Cells[i, 57] as Excel.Range).Value.ToString();
                        pracChe.Text = (xSheet.Cells[i, 58] as Excel.Range).Value.ToString();
                        tChe.Text = (xSheet.Cells[i, 59] as Excel.Range).Value.ToString();
                        gChe.Text = (xSheet.Cells[i, 60] as Excel.Range).Value.ToString();
                        gpaChe.Text = (xSheet.Cells[i, 61] as Excel.Range).Value.ToString();

                        //Biology
                        oBio.Text = (xSheet.Cells[i, 62] as Excel.Range).Value.ToString();
                        sBio.Text = (xSheet.Cells[i, 63] as Excel.Range).Value.ToString();
                        pracBio.Text = (xSheet.Cells[i, 64] as Excel.Range).Value.ToString();
                        tBio.Text = (xSheet.Cells[i, 65] as Excel.Range).Value.ToString();
                        gBio.Text = (xSheet.Cells[i, 66] as Excel.Range).Value.ToString();
                        gpaBio.Text = (xSheet.Cells[i, 67] as Excel.Range).Value.ToString();

                        //Higher Math
                        oHM.Text = (xSheet.Cells[i, 68] as Excel.Range).Value.ToString();
                        sHM.Text = (xSheet.Cells[i, 69] as Excel.Range).Value.ToString();
                        pracHM.Text = (xSheet.Cells[i, 70] as Excel.Range).Value.ToString();
                        tHM.Text = (xSheet.Cells[i, 71] as Excel.Range).Value.ToString();
                        gHM.Text = (xSheet.Cells[i, 72] as Excel.Range).Value.ToString();
                        gpaHM.Text = (xSheet.Cells[i, 73] as Excel.Range).Value.ToString();
                    }
                    else
                    {
                        clearLabels();
                    }
                }
                else
                {
                    status.Text = "nothing found";
                    clearLabels();

                }
                xApp.DisplayAlerts = false;
                xBook.Close();
                xApp.Quit();

            }
        }

        private void browse_Click(object sender, EventArgs e)
        {
            openToolStripMenuItem_Click(sender, e);
        }

        // on test (^_^) 
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(bmp,0, 0);
        }
        Bitmap bmp;
        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            Graphics g = this.CreateGraphics();
            bmp = new Bitmap(this.Size.Width, this.Size.Height, g);
            Graphics mg = Graphics.FromImage(bmp);
            mg.CopyFromScreen(this.Location.X, this.Location.Y, 0, 0, this.Size);
            if(printPreviewDialog1.ShowDialog()==DialogResult.OK)
            {
                printDocument1.Print();
            }
            
        }
    }
}
