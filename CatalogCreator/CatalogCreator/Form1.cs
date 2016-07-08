using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data.SqlClient;
using Core = Microsoft.Office.Core;

namespace CatalogCreator
{
    public partial class Form1 : Form
    {
        
        string path = Application.StartupPath+"/a.txt";
        public Form1()
        {
            InitializeComponent();
            
        }
        private void textBox1_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter="|*.csv";
            op.ShowDialog();
            textBox1.Text = op.FileName;
        }
        private void textBox2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog op = new FolderBrowserDialog();
            op.ShowDialog();
            textBox2.Text = op.SelectedPath;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //string[] contents = File.ReadAllLines(textBox1.Text);

            //List<_Class> classes = GetClass(contents);

            //CreateCatalog();
            // CreateCat();
            // CreateTableInDoc();
            //c();
            // CreateC();
            //CreateDocumentPropertyTable1();

           CreateDocumentPropertyTable();

           

        }
        
        private void CreateDocumentPropertyTable()
        {
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                string[] contents = File.ReadAllLines(textBox1.Text);
                string picturesPath = textBox2.Text+@"\";

                if (listView1.Items.Count > 0) { listView1.Items.Clear(); }
                if (listView1.Items.Count == 0)
                {

                    List<_Class> classes = GetClass(contents);

                    object objMiss = System.Reflection.Missing.Value;
                    object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                    Word._Application objApp;
                    Word._Document objDoc;
                    objApp = new Word.Application();
                    objApp.Visible = true;
                    objDoc = objApp.Documents.Add(ref objMiss, ref objMiss,
                       ref objMiss, ref objMiss);

                    objDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                    objDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                    objDoc.PageSetup.TopMargin = 20f;
                    objDoc.PageSetup.BottomMargin = 20f;
                    objDoc.PageSetup.LeftMargin = 20f;
                    objDoc.PageSetup.RightMargin = 20f;

                    //****************************************************************************************************//

                    for(int i=classes.Count-1;i>=0;i--)
                    {
                        object start = 0, end = 0;
                        Word.Range rng = objDoc.Range(ref start, ref end);

                        // Insert a title for the table and paragraph marks. 
                        rng.InsertBefore(classes[i].ClassName);
                        rng.Font.Name = "Times New Roman";
                        rng.Font.Size = 20;
                        rng.Font.Bold = 1;
                         
                        rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        rng.InsertParagraphAfter();
                        rng.InsertParagraphAfter();
                        rng.InsertParagraphAfter();

                        rng.SetRange(rng.Start, rng.End);

                        // Add the table.
                        rng.Tables.Add(objDoc.Paragraphs[2].Range, 1, 4, ref objMiss, ref objMiss);

                        // Format the table and apply a style. 
                        Word.Table tbl = rng.Tables[1];
                        tbl.Range.Font.Size = 12;
                        tbl.Columns.DistributeWidth();
                        tbl.Range.Paragraphs.SpaceAfter = 0;

                        object styleName = "Table Grid 1";
                        tbl.set_Style(ref styleName);

                        List<ItemsInfo> iteminfo = new List<ItemsInfo>();
                        iteminfo = classes[i].Items;
                        int row = 1, ii = 1;
                        bool isNew = false;
                        bool isHeight = false;
                        bool isWidth = false;
                        foreach (ItemsInfo it in iteminfo) {


                            //// Insert document properties into cells.   
                            tbl.Cell(row, ii).Range.Text = it.ItemNo + "\n" + it.Unit_Ctn + "\n" + it.Description;
                            Word.Range rnge = tbl.Cell(row, ii).Range.Paragraphs[1].Range;
                            rnge.Font.Size = 12;
                            rnge.Font.Bold = 1;
                            rnge.Font.Name = "Times New Roman";

                            Word.Range rnge2 = tbl.Cell(row, ii).Range.Paragraphs[2].Range;
                            rnge2.Font.Size = 12;
                            rnge2.Font.Bold = 0;
                            rnge2.Font.Name = "Times New Roman";

                            Word.Range rnge3 = tbl.Cell(row, ii).Range.Paragraphs[3].Range;
                            rnge3.Font.Size = 12;
                            rnge3.Font.Bold = 0;
                            rnge3.Font.Name = "Times New Roman";

                            string file = "";
                            bool isPng = false, isJpg = false,isJpeg=false;
                            try
                            {
                                byte[] b = File.ReadAllBytes(picturesPath + it.ItemNo + ".jpg");
                                isJpg = true;
                            }
                            catch { }

                            try
                            {
                                byte[] b = File.ReadAllBytes(picturesPath + it.ItemNo + ".jpeg");
                                isJpeg = true;
                            }
                            catch { }

                            try
                            {
                                byte[] bb = File.ReadAllBytes(picturesPath + it.ItemNo + ".png");
                                isPng = true;
                            }
                            catch { }

                            file = isJpg ? picturesPath + it.ItemNo + ".jpg" :
                                (isPng) ? picturesPath + it.ItemNo + ".png" : (isJpeg) ? picturesPath + it.ItemNo + ".jpeg" : ""; 

                            //file = @"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\55320.jpg";

                            if (file != "")
                            {
                                var shape = tbl.Cell(row, ii).Range.InlineShapes.AddPicture(file, ref objMiss, ref objMiss);

                                //shape1.LockAspectRatio = Core.MsoTriState.msoTrue;// Core.MsoTriState.msoTrue;
                                //shape1.ScaleHeight(0.5f, Core.MsoTriState.msoTrue,
                                //    Core.MsoScaleFrom.msoScaleFromTopLeft);

                                //// Align the picture on the upper right.
                                //if (File.Exists(Application.StartupPath + @"\b.csv"))
                                //{
                                //    File.AppendAllText(Application.StartupPath + @"\b.csv", ""+it.ItemNo +","+it.Description+","+shape.Width+","+shape.Height+ Environment.NewLine);
                                //}
                                //else
                                //{
                                //    File.AppendAllText(Application.StartupPath + @"\b.csv", "" + it.ItemNo + "," + it.Description + "," + shape.Width + "," + shape.Height + Environment.NewLine);
                                //}
                                //MessageBox.Show(shape.Height.ToString() + "  " + shape.Width.ToString());
                                shape.LockAspectRatio = Core.MsoTriState.msoCTrue;
                                //shape.Width = 160;
                                
                                //fix width
                                //fix height

                                if (!isNew)
                                {
                                    
                                    if (shape.Height > shape.Width) isHeight = true;
                                    else if (shape.Height < shape.Width) isWidth = true;
                                    isNew = true;
                                }

                                if (isHeight) shape.Height = 80;
                                else if (isWidth) shape.Width = 80;
                                //shape.Width = 105;
                                
                                //Word.Shape shape1 = shape.ConvertToShape();
                                //shape1.WrapFormat.Type = Word.WdWrapType.wdWrapThrough;

                                // shape1.LockAspectRatio = Core.MsoTriState.msoTrue;
                                // shape1.Height = 100;
                                // shape1.Width = 105;
                            }
                            else
                            {
                                if (File.Exists(Application.StartupPath + @"\c.csv"))
                                {
                                    File.AppendAllText(Application.StartupPath + @"\c.csv", "" + it.ItemNo + "," + it.Description + Environment.NewLine);
                                }
                                else
                                {
                                    File.AppendAllText(Application.StartupPath + @"\c.csv", "" + it.ItemNo + "," + it.Description + Environment.NewLine);
                                }
                            }
                            tbl.Cell(row, ii).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            //var shape = tbl.Cell(1, 1).Range.InlineShapes.AddPicture(@"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\01.jpg", ref objMiss, ref objMiss, ref objMiss);
                            ii++;
                            if (ii >4) {
                                
                               
                                if ((iteminfo.Count / 4 == row && iteminfo.Count % 4 == 0)) { } //dont add row in the last if count items is divisible by 3
                                else { tbl.Rows.Add(); }

                                ii = 1;
                                row++;
                            }
                        }
                         //*************************************************************************************************************//
                    }
                }
            }
        }
        public void CreateC()
        {
            Word._Application objApp;
            Word._Document objDoc;

            try
            {
                object objMiss = System.Reflection.Missing.Value;
                object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                //Start Word and create a new document.
                objApp = new Word.Application();
                objApp.Visible = true;
                objDoc = objApp.Documents.Add(ref objMiss, ref objMiss,
                    ref objMiss, ref objMiss);

                //Insert a paragraph at the end of the document.
                Word.Paragraph objPara2; //define paragraph object
                object oRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range; //go to end of the page
                objPara2 = objDoc.Content.Paragraphs.Add(ref oRng); //add paragraph at end of document
                
                objPara2.Format.SpaceAfter = 10; //defind some style
                objPara2.Range.InsertParagraphAfter(); //insert paragraph

                objDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                objDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperLetter;
                objDoc.PageSetup.TopMargin = 20f;
                objDoc.PageSetup.BottomMargin = 20f;
                objDoc.PageSetup.LeftMargin = 20f;
                objDoc.PageSetup.RightMargin = 20f;

                string[] contents = File.ReadAllLines(textBox1.Text);

                int ii = 1;

                int iRow, iCols = 0;
                int row = 2;
                string strText;

                List<_Class> classes = GetClass(contents);
                Word.Range objWordRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range; //go to end of document

                for (int k=classes.Count-1;k>=140; k--)
                {
                    Word.Range rng = objDoc.Range(0, 0);
                    Word.Table objTab1; //create table object
                    objTab1 = objDoc.Application.ActiveDocument.Tables.Add(rng, 2, 3, ref objMiss, ref objMiss); //add table object in word document
                    
                    //List<ItemsInfo> itms = new List<ItemsInfo>();
                    //itms = classes[k].Items;
                    //string classname = classes[k].ClassName;

                    //objTab1.Rows[1].Range.Text = classname;
                    //objTab1.Rows[1].Range.Font.Bold = 1;
                    //objTab1.Rows[1].Range.Font.Size = 24;
                    //objTab1.Rows[1].Range.Font.Name = "Times New Roman";
                    //objTab1.Rows[1].Cells[1].Merge(objTab1.Rows[1].Cells[3]);
                    

                    //foreach (ItemsInfo itm in itms)
                    //{
                    //    objTab1.Cell(row, ii).Range.Text = itm.ItemNo+"\n"+itm.Unit_Ctn+ "\n"+itm.Description; //add some text to cell
                    //    objTab1.Cell(row, ii).Range.Font.Name = "Times New Roman";
                    //    objTab1.Cell(row, ii).Range.Font.Size = 12;
                    //    objTab1.Cell(row, ii).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                         
                    //    ii++;
                    //    if (ii > 3) { ii = 0; row++; objTab1.Rows.Add(); objTab1.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow); }
                    //}

                    //object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    //object oPageBreak = Word.WdBreakType.wdPageBreak;
                    //objWordRng.Collapse(ref oCollapseEnd);
                    //objWordRng.InsertBreak(ref oPageBreak);
                    //objWordRng.Collapse(ref oCollapseEnd);
                }
                MessageBox.Show(objDoc.Application.ActiveDocument.Tables.Count.ToString());
                //for (iRow = 1; iRow <= contents.Length / 3; iRow++)
               // {
                    //if (contents[iRow].ToLower().Trim().Contains("item no")) { }
                    //else if (contents[iRow].ToLower().Contains("class")) { }
                    //else
                    //{

                    //    string[] t = contents[iRow].Split(new char[] { ',' });

                    //    strText = t[0] + "\n" + t[2] + "\n" + t[1];

                    //    string file = "";
                    //    bool isPng = false, isJpg = false;
                    //    try
                    //    {
                    //        byte[] b = File.ReadAllBytes(@"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".jpg");
                    //        isJpg = true;
                    //    }
                    //    catch { }

                    //    try
                    //    {
                    //        byte[] bb = File.ReadAllBytes(@"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".png");
                    //        isPng = true;
                    //    }
                    //    catch { }
                    //    file = isJpg ? @"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".jpg" :
                    //          (isPng) ? @"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".png" : "";


                    //    objTab1.Cell(row, ii).Range.Text = strText; //add some text to cell
                    //    objTab1.Cell(row, ii).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //    if (ii >= 1)
                    //    {

                    //        for (int i = 1; i < 4; i++)
                    //        {
                    //            Word.Range rng = objTab1.Cell(row, ii).Range.Paragraphs[i].Range;
                    //            // Change the formatting. To change the font size for a right-to-left language, 
                    //            // such as Arabic or Hebrew, use the Font.SizeBi property instead of Font.Size.
                    //            rng.Font.Size = 12;
                    //            if (i == 1) rng.Font.Bold = 1;
                    //            rng.Font.Name = "Times New Roman";
                    //            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    //        }
                    //    }

                    //    if (file != "") objTab1.Cell(row, ii).Range.InlineShapes.AddPicture(file, ref objMiss, ref objMiss, ref objMiss);
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    //    objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    //    objTab1.Rows.Height = 50f;
                    //    //objTab1.Rows[1].Range.Font.Bold = 1; //make first row of table BOLD
                    //    //objTab1.Columns[1].Width = objApp.InchesToPoints(3); //increase first column width
                    //    ii++;
                    //    if (ii > 3) { ii = 0; row++; }
                    //}
               // }

                ////Add some text after table
                //objWordRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range;
                //objWordRng.InsertParagraphAfter(); //put enter in document
                //objWordRng.InsertAfter("THIS IS THE SIMPLE WORD DEMO : THANKS YOU.");

                object szPath = "test.docx"; //your file gets saved with name 'test.docx'
                objDoc.SaveAs(ref szPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while executing code : " + ex.Message);
            }
            finally
            {
                //you can dispose object here
            }
        }
        public void c()
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Heading 1";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Heading 2";
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();

            //Insert another paragraph.
            Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 24;
            oPara3.Range.InsertParagraphAfter();

            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;
            string strText;
            for (r = 1; r <= 3; r++)
                for (c = 1; c <= 5; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;

            //Add some text after the table.
            Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "And here's another table:";
            oPara4.Format.SpaceAfter = 100;
            oPara4.Range.InsertParagraphBefore();

            //Insert a 5 x 2 table, fill it with data, and change the column widths.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            for (r = 1; r <= 5; r++)
                for (c = 1; c <= 2; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            oTable.Columns[2].Width = oWord.InchesToPoints(3);

            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            object oPos;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 6;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                               (Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));

            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;

            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
            wrdRng.Collapse(ref oCollapseEnd);

            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();

            //Insert a chart.
            Word.InlineShape oShape;
            object oClassType = "MSGraph.Chart.8";
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            object oChart;
            object oChartApp;
            oChart = oShape.OLEFormat.Object;
            oChartApp = oChart.GetType().InvokeMember("Application",
                BindingFlags.GetProperty, null, oChart, null);

            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
                null, oChart, Parameters);

            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update",
                BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit",
                BindingFlags.InvokeMethod, null, oChartApp, null);
            //... If desired, you can proceed from here using the Microsoft Graph 
            //Object model on the oChart and oChartApp objects to make additional
            //changes to the chart.

            //Set the width of the chart.
            oShape.Width = oWord.InchesToPoints(6.25f);
            oShape.Height = oWord.InchesToPoints(3.57f);

            //Add text after the chart.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("THE END.");

            //Close this form.
            this.Close();

        }
        public void CreateTableInDoc()
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Microsoft.Office.Interop.Word._Application objWord;
            Microsoft.Office.Interop.Word._Document objDoc;
            objWord = new Microsoft.Office.Interop.Word.Application();
            objWord.Visible = true;
            objDoc = objWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            int i = 0;
            int j = 0;

            Microsoft.Office.Interop.Word.Table objTable;
            Microsoft.Office.Interop.Word.Range wrdRng = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            string strText;
            objTable = objDoc.Tables.Add(wrdRng, 4, 2, ref oMissing, ref oMissing);
            objTable.Range.ParagraphFormat.SpaceAfter = 7;
            strText = "ABC Company";
            objTable.Rows[1].Range.Text = strText;
            objTable.Rows[1].Range.Font.Bold = 1;
            objTable.Rows[1].Range.Font.Size = 24;
            objTable.Rows[1].Range.Font.Position = 1;
            objTable.Rows[1].Range.Font.Name = "Times New Roman";
            objTable.Rows[1].Cells[1].Merge(objTable.Rows[1].Cells[2]);
            objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            objTable.Rows[2].Range.Font.Italic = 1;
            objTable.Rows[2].Range.Font.Size = 14;
            objTable.Cell(2, 1).Range.Text = "Item Name";
            objTable.Cell(2, 2).Range.Text = "Price";
            objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            objTable.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            for (i = 3; i <= 4; i++)
            {
                objTable.Rows[i].Range.Font.Bold = 0;
                objTable.Rows[i].Range.Font.Italic = 0;
                objTable.Rows[i].Range.Font.Size = 12;
                for (j = 1; j <= 2; j++)
                {
                    if (j == 1)
                        objTable.Cell(i, j).Range.Text = "Item " + (i - 1);
                    else
                        objTable.Cell(i, j).Range.Text = "Price of " + (i - 1);
                }
            }

            try
            {
                objTable.Borders.Shadow = true;
                objTable.Borders.Shadow = true;
            }
            catch
            {
            }

        }
        public void CreateCat()
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                string[] contents = File.ReadAllLines(textBox1.Text);

                if (listView1.Items.Count > 0) { listView1.Items.Clear(); }
                if (listView1.Items.Count == 0)
                {
                    List<_Class> classes = GetClass(contents);

                   
                    Word._Application objApp;
                    Word._Document objDoc;

                    try
                    {
                       
                        object objMiss = System.Reflection.Missing.Value;
                        object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                        //Start Word and create a new document.
                        objApp = new Word.Application();
                        objApp.Visible = true;
                        objDoc = objApp.Documents.Add(ref objMiss, ref objMiss,
                            ref objMiss, ref objMiss);

                        objDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                        objDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperLetter;
                        objDoc.PageSetup.TopMargin = 20f;
                        objDoc.PageSetup.BottomMargin = 20f;
                        objDoc.PageSetup.LeftMargin = 20f;
                        objDoc.PageSetup.RightMargin = 20f;

                        

                        foreach (_Class c in classes)
                        {
                             
                            List<ItemsInfo> itms = new List<ItemsInfo>();
                            itms = c.Items;
                            string classname = c.ClassName;
                            

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error occurred while executing code : " + ex.Message);
                    }
                    finally
                    {
                        //you can dispose object here
                    }

                }
            }

        }
        public class ItemsInfo
        {
            private string itemno;
            private string description;
            private string unit_ctn;

            public string ItemNo { get { return this.itemno; }set { this.itemno = value; } }
            public string Description { get { return this.description; }set { this.description = value; } }
            public string Unit_Ctn { get { return this.unit_ctn; }set { this.unit_ctn = value; } }
        }
        public class _Class
        {
            private List<ItemsInfo> items;
            private string classname;

            public string ClassName {get { return this.classname; }set { this.classname = value; } }
            public List<ItemsInfo> Items { get { return this.items; }set { this.items = value; } }
        }
        public List<_Class> GetClass(string[] contents)
        {
            int counterClass = 0;
            int counterClassItem = 0;
            string prev = "";
            string next = "";

            List<_Class> classes = new List<_Class>();
            List<ItemsInfo> itif = new List<ItemsInfo>();

            foreach (string s in contents)
            {

                if (s.ToLower().Contains("class") && s.ToLower().Contains("qty")) {

                    string[] t = s.Split(new char[] { ',' });

                    if (counterClass == 0) { prev = t[1]; next = ""; }
                    else { next = prev; prev = t[1]; }
                    
                    if (counterClass > 0)
                    {
                        _Class _class = new _Class();
                        _class.ClassName = next;
                        _class.Items = itif;

                        classes.Add(_class);

                        itif = new List<ItemsInfo>();

                        ListViewItem item = new ListViewItem(counterClass.ToString());
                        item.SubItems.Add(next);
                        item.SubItems.Add(counterClassItem.ToString());

                        listView1.Items.Add(item);

                    }
                    counterClass++;
                    counterClassItem = 0;
                }
                else if (s.ToLower().Trim().Contains("item no")) { }
                else {
                    string[] t = s.Split(new char[] { ',' });

                    ItemsInfo iit = new ItemsInfo();
                    iit.ItemNo = t[0];
                    iit.Description = t[1];
                    iit.Unit_Ctn = t[2];

                    itif.Add(iit);
                    counterClassItem++;
                }
            }

            _Class _classs = new _Class();
            _classs.ClassName = prev;
            _classs.Items = itif;
            classes.Add(_classs);

            ListViewItem itm = new ListViewItem(counterClass.ToString());
            itm.SubItems.Add(prev);
            itm.SubItems.Add(counterClassItem.ToString());

            listView1.Items.Add(itm);
           
            double it = 0;
            foreach(ListViewItem ii in listView1.Items)
            {
                it +=Convert.ToDouble(ii.SubItems[2].Text);
            }
            label3.Text = "Total items found per class : " + it;

            return classes;
        }
        public void CreateCatalog()
        {
            Word._Application objApp;
            Word._Document objDoc;

            try
            {
                object objMiss = System.Reflection.Missing.Value;
                object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                //Start Word and create a new document.
                objApp = new Word.Application();
                objApp.Visible = true;
                objDoc = objApp.Documents.Add(ref objMiss, ref objMiss,
                    ref objMiss, ref objMiss);

                //Insert a paragraph at the end of the document.
                Word.Paragraph objPara2; //define paragraph object
                object oRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range; //go to end of the page
                objPara2 = objDoc.Content.Paragraphs.Add(ref oRng); //add paragraph at end of document
                objPara2.Range.Text = "Test Table Caption"; //add some text in paragraph
                objPara2.Format.SpaceAfter = 10; //defind some style
                objPara2.Range.InsertParagraphAfter(); //insert paragraph

                objDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                objDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperLetter;
                objDoc.PageSetup.TopMargin = 20f;
                objDoc.PageSetup.BottomMargin = 20f;
                objDoc.PageSetup.LeftMargin = 20f;
                objDoc.PageSetup.RightMargin = 20f;

                string[] contents = File.ReadAllLines(textBox1.Text);

                //Insert a 2 x 2 table, (table with 2 row and 2 column)

                Word.Table objTab1; //create table object
                Word.Range objWordRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range; //go to end of document
                objTab1 = objDoc.Tables.Add(objWordRng, contents.Length / 3, 3, ref objMiss, ref objMiss); //add table object in word document

                objTab1.Range.ParagraphFormat.SpaceAfter = 2;

                int ii = 1;

                int iRow, iCols = 0;
                int row = 1;
                string strText;
                for (iRow = 1; iRow <= contents.Length / 3; iRow++)
                {

                    if (contents[iRow].ToLower().Trim().Contains("item no")) { }
                    else if (contents[iRow].ToLower().Contains("class")) { }
                    else
                    {

                        string[] t = contents[iRow].Split(new char[] { ',' });

                        strText = t[0] + "\n" + t[2] + "\n" + t[1];

                        string file = "";
                        bool isPng = false, isJpg = false;
                        try
                        {
                            byte[] b = File.ReadAllBytes(@"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".jpg");
                            isJpg = true;
                        }
                        catch { }

                        try
                        {
                            byte[] bb = File.ReadAllBytes(@"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".png");
                            isPng = true;
                        }
                        catch { }
                        file = isJpg ? @"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".jpg" :
                              (isPng) ? @"C:\Users\HUG\Desktop\~~~~ A R G I E   - F I L E S\CATALOG\PICTURES\" + t[0] + ".png" : "";


                        objTab1.Cell(row, ii).Range.Text = strText; //add some text to cell
                        objTab1.Cell(row, ii).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        if (ii >= 1)
                        {

                            for (int i = 1; i < 4; i++)
                            {
                                Word.Range rng = objTab1.Cell(row, ii).Range.Paragraphs[i].Range;
                                // Change the formatting. To change the font size for a right-to-left language, 
                                // such as Arabic or Hebrew, use the Font.SizeBi property instead of Font.Size.
                                rng.Font.Size = 12;
                                if (i == 1) rng.Font.Bold = 1;
                                rng.Font.Name = "Times New Roman";
                                rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            }
                        }

                        if (file != "") objTab1.Cell(row, ii).Range.InlineShapes.AddPicture(file, ref objMiss, ref objMiss, ref objMiss);
                        objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        objTab1.Cell(row, ii).Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        objTab1.Rows.Height = 50f;
                        //objTab1.Rows[1].Range.Font.Bold = 1; //make first row of table BOLD
                        //objTab1.Columns[1].Width = objApp.InchesToPoints(3); //increase first column width
                        ii++;
                        if (ii > 3) { ii = 0; row++; }
                    }

                }

                ////Add some text after table
                objWordRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range;
                objWordRng.InsertParagraphAfter(); //put enter in document
                objWordRng.InsertAfter("THIS IS THE SIMPLE WORD DEMO : THANKS YOU.");

                object szPath = "test.docx"; //your file gets saved with name 'test.docx'
                objDoc.SaveAs(ref szPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while executing code : " + ex.Message);
            }
            finally
            {
                //you can dispose object here
            }

        }
    }
}
