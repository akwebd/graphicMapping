using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Mapping
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Please select operation: 1-map, 2-restore, 3-update");
            int selection;
            selection = Convert.ToInt32(Console.ReadLine());
            //aux instances
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result;
            switch (selection)
            {
                case 1:
                    //mapping selected
                    openFileDialog1.InitialDirectory = "c:\\";
                    openFileDialog1.Title = "Select excel file to process";
                    openFileDialog1.Filter = "select excel with data (*.*)|*.*";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        try {
                            Mapping(openFileDialog1.FileName);
                        }                        
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                        }
                    }
                    //Mapping(@"C:\Users\DELL\Desktop\for graphic remaping\For graphics remaping\POINTS.xlsx");
                    break;
                case 2:
                    //restore selected
                    fbd.Description = "Select directory with files to restore";
                    result = fbd.ShowDialog();
                    if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        try
                        {
                            rewind(fbd.SelectedPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                        }
                    }
                    //rewind(@"C:\Users\DELL\Desktop\for graphic remaping\For graphics remaping\HMI_NEW");
                    break;
                case 3:
                    //update selected
                    openFileDialog1.InitialDirectory = "c:\\";
                    openFileDialog1.Title = "Select excel file to use for update";
                    openFileDialog1.Filter = "select excel with data (*.*)|*.*";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            Console.WriteLine("Excel table selected");
                            update(openFileDialog1.FileName);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                        }
                    }
                    break;
                default:
                    Console.WriteLine("nothing");
                    break;
            }
            //GetDATA(@"C:\Users\DELL\Desktop\for graphic remaping\For graphics remaping\POINTS.xlsx");
        }


        static void Mapping(string xls)
        {
            int lineL = 132; //source file line length limitation
            string[] colLetter = Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => ((Char)i).ToString()).ToArray();//GENERATE ALPHABET
            string directoryXLS = "", directorySRC = "", directoryNEW = "", directoryLOG = "", directoryTemp = ""; //explorer directory addresses
            string[] Files = null;//graphic SRC file addresses
            string[] points = null;//all points
            Queue linesOrig = new Queue();//to copy text lines
            Queue linesChanged = new Queue();//to store changed lines
            string[] macroses = null;//list of macroses to look at
            string[] diagrams = null;//list of diagrams to process
            string[] beginings = null; //list of possible beginings
            bool process2 = false;//when set of lines need to be processed
            bool process1 = false;//when new line is reached
            bool process3 = false;//when all necessarry data is collected
            string[] temp = null; //array to store line elements, separated by space
            int numL = 0;//used for line processing
            int numF = 0;
            int numTxt = 0; //number of text (FG) items
            int numConst = 0;//number of constants
            int numSet = 0;//number of SETs
            string[] tempLine = new string[3];//to store temporary lines
            Queue entries = new Queue();//to save place holde locations
            
            double rown = 0;
            string[] xtrText = { "NONE", "RTL", "TTB", "BOTH" };

            //Get excel directory
            directoryXLS = xls;

            //Get list of all SRC files
            directorySRC = directoryXLS.Substring(0, directoryXLS.LastIndexOf("\\")) + "\\HMI\\"; //graphics directory
            directoryNEW = directoryXLS.Substring(0, directoryXLS.LastIndexOf("\\")) + "\\HMI_NEW\\"; //result directory
            Files = Directory.GetFiles(directorySRC, "*.src", SearchOption.AllDirectories);

            ////Excel data acquisition
            Console.WriteLine("Excel data acquisition");
            Excel.Application app = new Excel.Application();
            Excel._Worksheet sheet;
            app.Visible = false;
            Excel.Workbook data = app.Workbooks.Open(@directoryXLS);
            try
            {
                sheet = data.Worksheets["DATA"];
            }
            catch
            {
                sheet = data.Worksheets.Add();
                sheet.Name = "DATA";
            }
            
            for (int cols = 1; cols < 5; cols++)
            {
                temp = null;
                rown = 10000 - 1 - (int)app.WorksheetFunction.CountBlank(sheet.get_Range(colLetter[cols - 1] + "2:" + colLetter[cols - 1] + "10000"));
                temp = new string[(int)rown];
                for (int i = 2; i < rown + 2; i++)
                {
                    temp[i - 2] = ((Excel.Range)sheet.Cells[i, cols]).Value;
                    temp[i - 2] = temp[i - 2].ToUpper();
                }

                switch (cols)
                {
                    case 1:
                        macroses = new string[temp.Length];
                        macroses = temp;
                        break;
                    case 2:
                        diagrams = new string[temp.Length];
                        diagrams = temp;
                        break;
                    case 3:
                        points = new string[temp.Length];
                        points = temp;
                        break;
                    case 4:
                        beginings = new string[temp.Length];
                        beginings = temp;
                        break;
                }
            }
            
            data.Close(SaveChanges: false, Filename: directoryXLS);
            app.Quit();
            Console.WriteLine("End of Excel data acquisition");
            ///end of data acquisition
            ///
            ///create new directory for mapping results

            if (!Directory.Exists(directoryNEW))
                Directory.CreateDirectory(directoryNEW);
            else
            {
                Files = Directory.GetFiles(directoryNEW, "*.src", SearchOption.AllDirectories);
                //remove all old *.SRC files
                foreach (string fl in Files)
                    File.Delete(fl);
            }

            Files = null;
            //file to log changes
            directoryLOG = directorySRC.Replace("HMI", "LOG");

            if (!Directory.Exists(directoryLOG))
                Directory.CreateDirectory(directoryLOG);
            else
            {
                Files = Directory.GetFiles(directoryLOG, "*.src", SearchOption.AllDirectories);
                //remove all old LOG  *.SRC files
                foreach (string fl in Files)
                    File.Delete(fl);
            }

            Files = Directory.GetFiles(directorySRC, "*.src", SearchOption.AllDirectories);
            ///data processing

            foreach (string str in Files)
            {
                directoryTemp = str.Replace("\\HMI\\", "\\HMI_NEW\\");
                //create HMI_NEW subfolders
                if (!Directory.Exists(directoryTemp.Substring(0, directoryTemp.LastIndexOf('\\'))))
                    Directory.CreateDirectory(directoryTemp.Substring(0, directoryTemp.LastIndexOf('\\')));
                numF++;
                Console.WriteLine(numF.ToString() + " in " + Files.Length.ToString());
                if (diagrams.Contains(str.ToUpper().Substring(str.LastIndexOf('\\') + 1)))
                {
                    directoryTemp = str.Replace("\\HMI\\", "\\LOG\\");
                    //create LOG subfolders
                    if (!Directory.Exists(directoryTemp.Substring(0, directoryTemp.LastIndexOf('\\'))))
                        Directory.CreateDirectory(directoryTemp.Substring(0, directoryTemp.LastIndexOf('\\')));
                    //logFile=File.AppendText(str.Replace("\\HMI\\", "\\LOG\\"));
                    //diagFile = File.AppendText(str.Replace("\\HMI\\", "\\HMI_NEW\\"));
                    foreach (string line in File.ReadLines(@str, Encoding.GetEncoding(1252)))
                    {
                        if ((line.Length < 1 || line.StartsWith("*")))
                            process1 = true;
                        else
                            if (!(line.Substring(0, 1) == " "))
                                process1=true;

                            //clear all stored data with each "new line"
                            if (process1)
                            {
                                process1=false;
                                //process if any line is stored
                                if (linesOrig.Count > 0)
                                {
                                    //check all stored lines for SRC points,
                                    //if found set the process point
                                    foreach (string line1 in linesOrig)
                                    {
                                        temp = null;
                                        tempLine[0]="";
                                        //split and check each line element
                                        temp = line1.ToUpper().Split(' ');
                                        foreach (string word in temp)
                                        {
                                            if(word.Contains("\\"))
                                            {
                                                tempLine[0] = word.Substring(word.IndexOf('\\') + 1, word.Length - word.IndexOf('\\') - 2);
                                                if (points.Contains(tempLine[0]))
                                                {
                                                    process2 = true;
                                                    break;
                                                }
                                            }                                            
                                        }

                                        if (process2)
                                            break;
                                    }
                                    //do mapping if SRC points were found
                                    if (process2)
                                    {
                                        process3 = true;
                                        numL = 0;
                                        process2 = false;
                                        temp = null;
                                        ///get macro number and replace macro
                                        temp = linesOrig.Peek().ToString().Split(' ');
                                        if (temp[0].ToUpper() == "MACRO")
                                        {
                                            //process line by line
                                            foreach (string line2 in linesOrig)
                                            {
                                                //swap macroses and parameters
                                                if (numL < 1)
                                                {
                                                    tempLine[0] = "";
                                                    //create the first line of macro command
                                                    //assuming that parameters of a macro fit into a first line
                                                    try
                                                    {
                                                        tempLine[1] = line2.Substring(line2.IndexOf('\\'));
                                                        tempLine[1] = tempLine[1].Replace("\\", "\"");
                                                    }
                                                    catch
                                                    {
                                                        tempLine[1] = "";
                                                    }
                                                    //get macro parameter
                                                    if (xtrText.Contains(temp[6]))
                                                    {
                                                        //how many text and constant elements there should be
                                                        numTxt = Convert.ToInt16(temp[10]) + Convert.ToInt16(temp[11])
                                                                + Convert.ToInt16(temp[12]);
                                                        numConst = Convert.ToInt16(temp[14]);
                                                        numSet = Convert.ToInt16(temp[13]);

                                                        tempLine[0] = temp[0] + " "
                                                            + temp[1] + "_STATIC "
                                                            + temp[2] + " "
                                                            + temp[3] + " "
                                                            + temp[4] + " "
                                                            + temp[5] + " "
                                                            + temp[6] + " "
                                                            + temp[7] + " "
                                                            + temp[8] + " "
                                                            + temp[9] + " 0 "
                                                            + numTxt.ToString()
                                                            + " 0 0 "
                                                            + temp[14] + " 0 ";

                                                    }
                                                    else
                                                    {
                                                        //how many text and constant elements there should be
                                                        numTxt = Convert.ToInt16(temp[6]) + Convert.ToInt16(temp[7])
                                                            + Convert.ToInt16(temp[8]);
                                                        numConst = Convert.ToInt16(temp[10]);
                                                        numSet = Convert.ToInt16(temp[9]);

                                                        tempLine[0] = temp[0] + " "
                                                            + temp[1] + "_STATIC "
                                                            + temp[2] + " "
                                                            + temp[3] + " "
                                                            + temp[4] + " "
                                                            + temp[5] + " 0 "
                                                            + numTxt.ToString()
                                                            + " 0 0 "
                                                            + temp[10] + " 0 ";
                                                    }
                                                    numL++;
                                                }
                                                else
                                                {
                                                    //convert all points to texts                                                    
                                                    try
                                                    {
                                                        tempLine[1] = line2.Replace("\\", "\"");
                                                    }
                                                    catch
                                                    {
                                                        tempLine[1] = line2;
                                                    }
                                                    tempLine[0] = "   ";
                                                    tempLine[1] = tempLine[1].Substring(3);

                                                }
                                                //store correct numer of variables
                                                while (tempLine[1].Length > 0 && process3)
                                                {
                                                    //end with no colors defined
                                                    if (numTxt < 1 && numSet < 1 && numConst < 1)
                                                    {
                                                        tempLine[0] = tempLine[0] + "0 0";
                                                        tempLine[1] = "";
                                                        process3 = false;
                                                        break;
                                                    }

                                                    //copy all CONST numbers
                                                    if (numTxt < 1 && numSet < 1) 
                                                    {
                                                        //to avoid hitting end of string
                                                        try
                                                        {
                                                            tempLine[0] = tempLine[0] + tempLine[1].Substring(0, tempLine[1].IndexOf(" ")) + " ";

                                                            numConst--;
                                                            tempLine[1] = tempLine[1].Substring(tempLine[1].IndexOf(" ") + 1);
                                                        }
                                                        catch
                                                        {
                                                            tempLine[1] = "";
                                                        }
                                                    }
                                                    //dissmiss all SET numbers
                                                    if (numTxt < 1 && numSet > 0)
                                                    {
                                                        tempLine[1] = tempLine[1].Substring(tempLine[1].IndexOf(" ") + 1);
                                                        numSet--;
                                                    }

                                                    //coppy all texts
                                                    if (numTxt > 0)
                                                    {

                                                        //to avoid hitting end of string
                                                        try
                                                        {
                                                            tempLine[0] = tempLine[0] + tempLine[1].Substring(0, tempLine[1].Substring(1).IndexOf("\"") + 2) + " ";
                                                            numTxt--;
                                                            tempLine[1] = tempLine[1].Substring(tempLine[1].Substring(1).IndexOf("\"") + 3);
                                                        }
                                                        catch
                                                        {
                                                            tempLine[1] = "";
                                                        }
                                                        //Console.WriteLine(tempLine[0]);
                                                        //Console.WriteLine(tempLine[1]);
                                                       //Console.WriteLine("cikle1");
                                                        
                                                    }
                                                }
                                                linesChanged.Enqueue(tempLine[0]);
                                            }
                                            
                                        }
                                        else
                                        {
                                            //process line by line, comment them
                                            foreach (string line3 in linesOrig)
                                                linesChanged.Enqueue("*" + line3);
                                        }
                                    //store the results to log and diag files+
                                    File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "#from:\r\n", Encoding.GetEncoding(1252));
                                        foreach (string line4 in linesOrig)
                                            File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), line4+"\r\n", Encoding.GetEncoding(1252));
                           
                                    File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "#to:\r\n", Encoding.GetEncoding(1252));
                                        //check line length before storing it
                                        numL = 0;
                                        temp = null;
                                        temp = new string[linesChanged.Count];
                                        linesChanged.CopyTo(temp, 0);
                                        //store text line and check that they are not too long
                                        foreach (string line4 in temp)
                                        {
                                            tempLine[0] = line4;
                                            while (tempLine[0].Length > lineL)
                                                {
                                                    tempLine[0] = tempLine[0].Substring(0, tempLine[0].Length - tempLine[0].LastIndexOf(' ') + 1);
                                                    try
                                                    {
                                                        temp[numL + 1] = "   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1) + temp[numL + 1];
                                                    }
                                                    catch
                                                    {
                                                        File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1)+"\r\n", Encoding.GetEncoding(1252));
                                                        File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), "   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1)+"\r\n", Encoding.GetEncoding(1252));
                                                        //logFile.WriteLine("   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1));
                                                        //diagFile.WriteLine("   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1));
                                                    }
                                                }

                                            File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), tempLine[0]+"\r\n", Encoding.GetEncoding(1252));
                                            File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), tempLine[0]+"\r\n", Encoding.GetEncoding(1252));
                                        
                                        //logFile.WriteLine(tempLine[0]);
                                        //diagFile.WriteLine(tempLine[0]);
                                        numL++;
                                        }
                                        File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "\r\n", Encoding.GetEncoding(1252));
                                    }
                                    else
                                    {
                                        //store original text lines
                                        foreach (string line4 in linesOrig)
                                            File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), line4+"\r\n", Encoding.GetEncoding(1252));
                                }
                                }
                                linesOrig.Clear();//clear queue with each primary line
                                linesChanged.Clear();//clear queue with each primary line
                            }
                            //copy empty or commented line
                            if (line.Length < 0 || line.StartsWith("*"))
                                File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), line+"\r\n", Encoding.GetEncoding(1252));
                            else
                                linesOrig.Enqueue(line);//queue all lines from a single command
                    }
                    //process if any line is stored
                    if (linesOrig.Count > 0)
                    {
                        //check all stored lines for SRC points,
                        //if found set the process point
                        foreach (string line1 in linesOrig)
                        {
                            temp = null;
                            tempLine[0] = "";
                            //split and check each line element
                            temp = line1.ToUpper().Split(' ');
                            foreach (string word in temp)
                            {
                                if (word.Contains("\\"))
                                {
                                    tempLine[0] = word.Substring(word.IndexOf('\\') + 1, word.Length - word.IndexOf('\\') - 2);
                                    if (points.Contains(tempLine[0]))
                                    {
                                        process2 = true;
                                        break;
                                    }
                                }
                            }

                            if (process2)
                                break;
                        }
                        //do mapping if SRC points were found
                        if (process2)
                        {
                            process3 = true;
                            numL = 0;
                            process2 = false;
                            temp = null;
                            ///get macro number and replace macro
                            temp = linesOrig.Peek().ToString().Split(' ');
                            if (temp[0].ToUpper() == "MACRO")
                            {
                                //process line by line
                                foreach (string line2 in linesOrig)
                                {
                                    //swap macroses and parameters
                                    if (numL < 1)
                                    {
                                        tempLine[0] = "";
                                        //create the first line of macro command
                                        //assuming that parameters of a macro fit into a first line
                                        try
                                        {
                                            tempLine[1] = line2.Substring(line2.IndexOf('\\'));
                                            tempLine[1] = tempLine[1].Replace("\\", "\"");
                                        }
                                        catch
                                        {
                                            tempLine[1] = "";
                                        }
                                        //get macro parameter
                                        if (xtrText.Contains(temp[6]))
                                        {
                                            //how many text and constant elements there should be
                                            numTxt = Convert.ToInt16(temp[10]) + Convert.ToInt16(temp[11])
                                                    + Convert.ToInt16(temp[12]);
                                            numConst = Convert.ToInt16(temp[14]);
                                            numSet = Convert.ToInt16(temp[13]);

                                            tempLine[0] = temp[0] + " "
                                                + temp[1] + "_STATIC "
                                                + temp[2] + " "
                                                + temp[3] + " "
                                                + temp[4] + " "
                                                + temp[5] + " "
                                                + temp[6] + " "
                                                + temp[7] + " "
                                                + temp[8] + " "
                                                + temp[9] + " 0 "
                                                + numTxt.ToString()
                                                + " 0 0 "
                                                + temp[14] + " 0 ";

                                        }
                                        else
                                        {
                                            //how many text and constant elements there should be
                                            numTxt = Convert.ToInt16(temp[6]) + Convert.ToInt16(temp[7])
                                                + Convert.ToInt16(temp[8]);
                                            numConst = Convert.ToInt16(temp[10]);
                                            numSet = Convert.ToInt16(temp[9]);

                                            tempLine[0] = temp[0] + " "
                                                + temp[1] + "_STATIC "
                                                + temp[2] + " "
                                                + temp[3] + " "
                                                + temp[4] + " "
                                                + temp[5] + " 0 "
                                                + numTxt.ToString()
                                                + " 0 0 "
                                                + temp[10] + " 0 ";
                                        }
                                        numL++;
                                    }
                                    else
                                    {
                                        //convert all points to texts                                                    
                                        try
                                        {
                                            tempLine[1] = line2.Replace("\\", "\"");
                                        }
                                        catch
                                        {
                                            tempLine[1] = line2;
                                        }
                                        tempLine[0] = "   ";
                                        tempLine[1] = tempLine[1].Substring(3);

                                    }
                                    //store correct numer of variables
                                    while (tempLine[1].Length > 0 && process3)
                                    {
                                        //end with no colors defined
                                        if (numTxt < 1 && numSet < 1 && numConst < 1)
                                        {
                                            tempLine[0] = tempLine[0] + "0 0";
                                            tempLine[1] = "";
                                            process3 = false;
                                            break;
                                        }

                                        //copy all CONST numbers
                                        if (numTxt < 1 && numSet < 1)
                                        {
                                            //to avoid hitting end of string
                                            try
                                            {
                                                tempLine[0] = tempLine[0] + tempLine[1].Substring(0, tempLine[1].IndexOf(" ")) + " ";

                                                numConst--;
                                                tempLine[1] = tempLine[1].Substring(tempLine[1].IndexOf(" ") + 1);
                                            }
                                            catch
                                            {
                                                tempLine[1] = "";
                                            }
                                        }
                                        //dissmiss all SET numbers
                                        if (numTxt < 1 && numSet > 0)
                                        {
                                            tempLine[1] = tempLine[1].Substring(tempLine[1].IndexOf(" ") + 1);
                                            numSet--;
                                        }

                                        //coppy all texts
                                        if (numTxt > 0)
                                        {

                                            //to avoid hitting end of string
                                            try
                                            {
                                                tempLine[0] = tempLine[0] + tempLine[1].Substring(0, tempLine[1].Substring(1).IndexOf("\"") + 2) + " ";
                                                numTxt--;
                                                tempLine[1] = tempLine[1].Substring(tempLine[1].Substring(1).IndexOf("\"") + 3);
                                            }
                                            catch
                                            {
                                                tempLine[1] = "";
                                            }
                                            //Console.WriteLine(tempLine[0]);
                                            //Console.WriteLine(tempLine[1]);
                                            //Console.WriteLine("cikle1");

                                        }
                                    }
                                    linesChanged.Enqueue(tempLine[0]);
                                }

                            }
                            else
                            {
                                //process line by line, comment them
                                foreach (string line3 in linesOrig)
                                    linesChanged.Enqueue("*" + line3);
                            }
                            //store the results to log and diag files+

                            File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "#from:\r\n", Encoding.GetEncoding(1252));
                            foreach (string line4 in linesOrig)
                                File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), line4 + "\r\n", Encoding.GetEncoding(1252));

                            File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "#to:\r\n", Encoding.GetEncoding(1252));
                            //check line length before storing it
                            numL = 0;
                            temp = null;
                            temp = new string[linesChanged.Count];
                            linesChanged.CopyTo(temp, 0);
                            //store text line and check that they are not too long
                            foreach (string line4 in temp)
                            {
                                tempLine[0] = line4;
                                while (tempLine[0].Length > lineL)
                                {
                                    tempLine[0] = tempLine[0].Substring(0, tempLine[0].Length - tempLine[0].LastIndexOf(' ') + 1);
                                    try
                                    {
                                        temp[numL + 1] = "   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1) + temp[numL + 1];
                                    }
                                    catch
                                    {
                                        File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1) + "\r\n", Encoding.GetEncoding(1252));
                                        File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), "   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1) + "\r\n", Encoding.GetEncoding(1252));
                                        //logFile.WriteLine("   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1));
                                        //diagFile.WriteLine("   " + tempLine[0].Substring(tempLine[0].LastIndexOf(' ') + 1));
                                    }
                                }
                                File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), tempLine[0] + "\r\n", Encoding.GetEncoding(1252));
                                File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), tempLine[0] + "\r\n", Encoding.GetEncoding(1252));
                                //logFile.WriteLine(tempLine[0]);
                                //diagFile.WriteLine(tempLine[0]);
                                numL++;
                            }
                            File.AppendAllText(str.Replace("\\HMI\\", "\\LOG\\"), "\r\n", Encoding.GetEncoding(1252));
                        }
                        else
                        {
                            //store original text lines
                            foreach (string line4 in linesOrig)
                                File.AppendAllText(str.Replace("\\HMI\\", "\\HMI_NEW\\"), line4 + "\r\n", Encoding.GetEncoding(1252));
                        }
                    }
                    linesOrig.Clear();//clear queue with each primary line
                    linesChanged.Clear();//clear queue with each primary line
                }
                //DO NOT COPY FILES
                //else//resave diagram in the new location
                //    File.Copy(str, str.Replace("\\HMI\\", "\\HMI_NEW\\"));
            }            
            ///end of datta proccessing

        }
        static void GetDATA(string xls)
        {
            string[] colLetter = Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => ((Char)i).ToString()).ToArray();//GENERATE ALPHABET
            string directoryXLS = "", directorySRC = ""; //explorer directory addresses
            string[] Files = null;//graphic SRC file addresses
            string[] points = null;//all points
            Stack macroses = new Stack();//used macroses
            Stack usedPoints = new Stack();//used points
            Stack Gdiagrams = new Stack();//points used in diagrams
            Stack beginings = new Stack();//beginings of text line
            string tempStr = "";
            string[] tempStr1; //tempStr used for point detection, tempStr1 used for line begining search
            Queue lines = new Queue();//to store text lines
            bool flag = false;
            int num = 0;

            //Get excel directory
            //Console.WriteLine("Excel address:");
            directoryXLS = xls;

            //Get list of all SRC files in that directory;
            directorySRC = directoryXLS.Substring(0, directoryXLS.LastIndexOf("\\")) + "\\HMI\\"; //graphics directory
            Files = Directory.GetFiles(directorySRC, "*.src", SearchOption.AllDirectories);

            ////Excel data acquisition
            Excel.Application app = new Excel.Application();
            app.Visible = false;

            Excel.Workbook data = app.Workbooks.Open(@directoryXLS);//open excell workbook
            Excel._Worksheet sheet = data.Worksheets["SRC"];//select excel spreadsheet

            double rown = 10000 - 1 - (int)app.WorksheetFunction.CountBlank(sheet.get_Range("A2:A10000"));//number of points
            points = new string[(int)rown];
            for (int i = 2; i < rown + 2; i++)
            {
                points[i - 2] = ((Excel.Range)sheet.Cells[i, 1]).Value;//read point names
                points[i - 2] = points[i - 2].ToUpper();
            }
            data.Close(SaveChanges: false, Filename: directoryXLS);//close excel file


            ////end of data acquisition

            ////collect information
            foreach (string str in Files)
            {
                num++;
                Console.WriteLine(num.ToString() + " is " + Files.Length);
                Console.WriteLine(str);
                foreach (string line in File.ReadLines(@str))
                {


                    flag = false;
                    if (!(line.IndexOf(' ') == 0))
                        lines.Clear();//clear queue when with each primary line

                    lines.Enqueue(line);
                    tempStr = line.ToUpper();
                    //Console.WriteLine(line);
                    if (tempStr.Contains('\\'))//if point identifier is found
                    {
                        tempStr = line.Substring(line.IndexOf('\\') + 1, line.Length - line.IndexOf('\\') - 1).ToUpper();
                        while (tempStr.Length > 0)
                        {
                            if (tempStr.Contains('\\'))//if point identifier is found
                            {
                                if (points.Contains(tempStr.Substring(0, tempStr.IndexOf('\\'))))
                                {

                                    //search of used points
                                    if (!usedPoints.Contains(tempStr.Substring(0, tempStr.IndexOf('\\'))))
                                        usedPoints.Push(tempStr.Substring(0, tempStr.IndexOf('\\')));

                                    if (!flag)
                                    {
                                        //store diagram numbers
                                        if (!Gdiagrams.Contains(str.Substring(str.LastIndexOf('\\') + 1, str.Length - str.LastIndexOf('\\') - 1)))
                                            Gdiagrams.Push(str.Substring(str.LastIndexOf('\\') + 1, str.Length - str.LastIndexOf('\\') - 1));

                                        //check for line begining
                                        if (!(line.IndexOf(' ') == 0))//if the line is primary
                                        {
                                            //search for used macro
                                            tempStr1 = line.ToUpper().Split(' ');
                                            if (tempStr1.Contains("MACRO") && !(line.Substring(0, 1) == "*"))
                                            {
                                                if (tempStr1.Length > 1)
                                                    if (!macroses.Contains(tempStr1[0] + " " + tempStr1[1]))
                                                        macroses.Push(tempStr1[0] + " " + tempStr1[1]);
                                            }
                                            flag = true;//begining for this line is found
                                            if (!beginings.Contains(line.Substring(0, line.IndexOf(' ')).ToUpper()))
                                            {
                                                beginings.Push(line.Substring(0, line.IndexOf(' ')).ToUpper()); //add begining type to the file
                                            }
                                            if (line.IndexOf(' ') == 0)//if the line is not primary
                                            {
                                                foreach (string stStr in lines)//seach in queue
                                                    if (!(stStr.IndexOf(' ') == 0))
                                                    {
                                                        //search for used macro
                                                        tempStr1 = stStr.ToUpper().Split(' ');
                                                        if (tempStr1.Contains("MACRO") && !(stStr.Substring(0, 1) == "*"))
                                                        {
                                                            if (tempStr1.Length > 1)
                                                                if (!macroses.Contains(tempStr1[0] + " " + tempStr1[1]))
                                                                    macroses.Push(tempStr1[0] + " " + tempStr1[1]);
                                                            break;
                                                        }
                                                        if (!beginings.Contains(stStr.Substring(0, stStr.IndexOf(' '))))
                                                        {
                                                            beginings.Push(stStr.Substring(0, stStr.IndexOf(' '))); //add begining type to the file
                                                        }
                                                    }
                                            }
                                        }
                                    }

                                    //move to the next point identifier
                                    if (tempStr.Contains('\\'))//if point identifier is found
                                        tempStr = tempStr.Substring(tempStr.IndexOf('\\') + 1, tempStr.Length - tempStr.IndexOf('\\') - 1);
                                    else
                                        break;

                                    if (tempStr.Contains('\\'))//if point identifier is found
                                        tempStr = tempStr.Substring(tempStr.IndexOf('\\') + 1, tempStr.Length - tempStr.IndexOf('\\') - 1);
                                    else
                                        break;
                                }
                                else
                                    break;
                            }
                            else
                                break;
                        }
                    }
                }
            }
            ////end of collect information

            ////store information
            data = app.Workbooks.Open(@directoryXLS);//open excell workbook
            Excel._Worksheet dataSh = null;
            //Store used macroses
            try
            {
                dataSh = data.Application.Worksheets["DATA"];
            }
            catch
            {
                dataSh=data.Application.Worksheets.Add();
                dataSh.Name="DATA";
            }

            rown = 0;
            //store used macroses
            Excel.Range range = dataSh.Cells[1, 1];
            range.Value = "Macroses";
            rown = 1;
            foreach (string txt in macroses)
            {
                rown++;
                range = dataSh.Cells[rown, 1];
                range.Value = txt;
            }


            //store used diagrams
            range = dataSh.Cells[1, 2];
            range.Value = "Diagrams";
            rown = 1;
            foreach (string txt in Gdiagrams)
            {
                rown++;
                range = dataSh.Cells[rown, 2];
                range.Value = txt;
            }

            //store used points
            range = dataSh.Cells[1, 3];
            range.Value = "Points";
            rown = 1;
            foreach (string txt in usedPoints)
            {
                rown++;
                range = dataSh.Cells[rown, 3];
                range.Value = txt;
            }

           /*//store beginings
            range = dataSh.Cells[1, 4];
            range.Value = "Begininings";
            rown = 1;
            foreach (string txt in beginings)
            {
                rown++;
                range = dataSh.Cells[rown, 4];
                range.Value = txt;
            }
            * */


            data.Close(SaveChanges: true, Filename: directoryXLS);//close excel file
            ////end of store information


        }

        static void rewind(string directory)
        {
            string directoryLOG = "", directorySRC_R = "", directoryNEW = ""; //explorer directory addresses
            string[] logFiles = null;//log file adrresses
            string[] diagFiles = null;//.SRC file adrresses
            Queue fromLines = new Queue();
            Queue toLines = new Queue();
            Queue logLines = new Queue();
            string tempLine = "";
            bool from = false, to = false, repl = false;
            string line1 = "";
            int modifications=0;
            string tempFile = "";//temporary file adrress
            string[] fromData=null;//to store coordinates and size
            string[] toData = null;//to store coordinates and size
            directoryNEW = directory;
            directoryLOG = directory.Replace("HMI_NEW", "LOG");
            directorySRC_R = directory.Replace("HMI_NEW", "HMI_RETURN");

            //clear destination directory

            if (!Directory.Exists(directorySRC_R))
                Directory.CreateDirectory(directorySRC_R);
            else
            {
                diagFiles = Directory.GetFiles(directorySRC_R, "*.src", SearchOption.AllDirectories);
                //remove all *.SRC files
                foreach (string fl in diagFiles)
                    File.Delete(fl);
            }
            diagFiles = null;

            //get lists of files
            diagFiles = Directory.GetFiles(directoryNEW, "*.src", SearchOption.AllDirectories);
            logFiles = Directory.GetFiles(directoryLOG, "*.src", SearchOption.AllDirectories);


            foreach (string file in diagFiles)
            {
                if (fromLines.Count > 0)
                {
                    Console.WriteLine("FROM");
                    while (fromLines.Count > 0)
                    {
                        
                        Console.WriteLine((string)fromLines.Dequeue());
                        modifications++;
                    }
                    Console.WriteLine(tempFile);
                }
                if (logLines.Count > 0)
                {
                    Console.WriteLine("LOADED");

                    while (logLines.Count > 0)
                    {                        
                        if((string)logLines.Peek()=="#from:")
                        Console.WriteLine((string)logLines.Dequeue());                        
                    }
                }
                to = false;
                from = false;
                toLines.Clear();
                toData = null;
                fromData = null;
                logLines.Clear();
                tempFile = file.Replace("\\HMI_NEW\\", "\\HMI_RETURN\\");
                //create HMI_NEW subfolders
                if (!Directory.Exists(tempFile.Substring(0, tempFile.LastIndexOf('\\'))))
                    Directory.CreateDirectory(tempFile.Substring(0, tempFile.LastIndexOf('\\')));
                tempFile = file.Replace("HMI_NEW", "LOG");
                if (logFiles.Contains(tempFile))
                {
                                      
                    //read all logged data
                    foreach (string line in File.ReadLines(@file.Replace("HMI_NEW", "LOG"), Encoding.GetEncoding(1252)))
                        logLines.Enqueue(line);

                    //file for storage
                    //diagFile = File.AppendText(file.Replace("\\HMI_NEW\\", "\\HMI_RETURN\\"));
                    foreach (string line in File.ReadLines(@file, Encoding.GetEncoding(1252)))
                    {
                        line1 = line;
                            //copy data to replace
                            if (fromLines.Count == 0)
                            {
                                repl = false;
                                while (logLines.Count > 0)
                                {
                                    // get line by line
                                    tempLine = (string)logLines.Dequeue();
                                    if (tempLine.Length != 0 || (!from && !to))
                                    {
                                        if (tempLine.StartsWith("Macro") || tempLine.StartsWith("MACRO"))//get macro parameters X Y W H
                                            fromData=tempLine.Split(' ');

                                        if (from)
                                            fromLines.Enqueue(tempLine);

                                        if (!from && to && tempLine == "#to:")
                                        {
                                            from = true;
                                            to = false;
                                        }

                                        if (to)
                                            toLines.Enqueue(tempLine);

                                        if (!to && tempLine == "#from:")
                                            to = true;
                                    }
                                    else
                                        if (fromLines.Count > 0)
                                        {
                                            from = false;
                                            break;
                                        }

                                }
                            }
                            try
                            {
                                toData=null;
                                toData = line1.Split(' ');
                                if (toData[0].StartsWith("*") && line1 == (string)fromLines.Peek())//for commented line replacement
                                    {
                                        fromLines.Dequeue();
                                        line1 = (string)toLines.Dequeue();
                                    }
                                else
                                    if(fromData.Length>5)
                                        if (repl || (fromData[2] == toData[2] && fromData[3] == toData[3] && fromData[4] == toData[4] && fromData[5] == toData[5]))//for macro replacemen
                                        {
                                            repl = true;
                                            fromLines.Dequeue();
                                            line1 = (string)toLines.Dequeue();
                                        }

                            }
                            catch { }
                            tempFile = file.Replace("\\HMI_NEW\\", "\\HMI_RETURN\\");
                            File.AppendAllText(tempFile, line1 + "\r\n", Encoding.GetEncoding(1252));
                    }
                }
                //DO NOT COPY FILES
                //else
                    //copy file if it was not previously processed
                    //File.Copy(file, file.Replace("\\HMI_NEW\\", "\\HMI_RETURN\\"));
            }
            Console.WriteLine(modifications);
        }
        static void update(string xls)
        {
            string[] colLetter = Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => ((Char)i).ToString()).ToArray();//GENERATE ALPHABET
            string directoryXLS = "", directoryUPD = "", directoryNEW = "", directoryLOG = "", directoryLOGUPD = "", dataFile = ""; //explorer directory addresses
            string[] diagFiles = null, logFiles = null;
            string[] pointsFrom = null;//all original points
            string[] pointsTo = null;//all new points
            string[] dataTo = null;//to store coordinates and size
            string[] dataFrom = null;//to store coordinates and size
            string actLine = "", tempLine="", tempLine1 = "";//string for data pocessing
            Queue linesToOld = new Queue();//original TO line
            Queue linesToNew = new Queue();//modified TO line
            Queue linesChanged = new Queue();//to store changed lines
            Queue linesFrom = new Queue();//original line
            Queue linesLog = new Queue();
            Queue linesLogUPD = new Queue();
            bool readTo, readFrom, procReplace=false, chngLog;
            int modifications = 0; //count total modifications
            string[] temp = null; //array to store line elements, separated by space
            Queue entries = new Queue();//to save place holde locations
            double rown = 0;
            string[] xtrText = { "NONE", "RTL", "TTB", "BOTH" };
            string[] tempQueue = new string[3];

            ////Excel data acquisition
            //Get excel directory
            directoryXLS = xls;
            Console.WriteLine("Excel data acquisition");
            Excel.Application app = new Excel.Application();
            Excel._Worksheet sheet=null;
            app.Visible = false;
            Excel.Workbook data = app.Workbooks.Open(@directoryXLS);
            try
            {
                sheet = data.Worksheets["POINTS"];
            }
            catch (Exception ex)
            {
                MessageBox.Show("POINTS spreadsheet does not exists. Original error: " + ex.Message);
            }
            //read OLD and NEW points
            for (int cols = 1; cols < 3; cols++)
            {
                temp = null;
                rown = 10000 - 1 - (int)app.WorksheetFunction.CountBlank(sheet.get_Range(colLetter[cols - 1] + "2:" + colLetter[cols - 1] + "10000"));
                temp = new string[(int)rown];
                for (int i = 2; i < rown + 2; i++)
                {
                    temp[i - 2] = ((Excel.Range)sheet.Cells[i, cols]).Value;
                    temp[i - 2] = temp[i - 2].ToUpper();
                }

                switch (cols)
                {
                    case 1:
                        pointsFrom = new string[temp.Length];
                        pointsFrom = temp;
                        break;
                    case 2:
                        pointsTo = new string[temp.Length];
                        pointsTo = temp;
                        break;
                }
            }
            data.Close(SaveChanges: false, Filename: directoryXLS);
            app.Quit();
            Console.WriteLine("End of Excel data acquisition");
            //end of data acquisition
            //start procesign
            //get data directories check if they exist,create if not
            directoryNEW = directoryXLS.Substring(0, directoryXLS.LastIndexOf("\\")) + "\\HMI_NEW\\"; //input file directory
            directoryUPD = directoryNEW.Replace("HMI_NEW", "HMI_UPDATE");//modified diagram directory
            directoryLOG = directoryNEW.Replace("HMI_NEW", "LOG");//modified diagram directory
            directoryLOGUPD = directoryNEW.Replace("HMI_NEW", "LOG_UPDATE");//modified diagram directory
            //get lists of files
            diagFiles = Directory.GetFiles(directoryNEW, "*.src", SearchOption.AllDirectories);
            logFiles = Directory.GetFiles(directoryLOG, "*.src", SearchOption.AllDirectories);

            //loop through all graphic diagrams and log files
            foreach (string logFile in logFiles)
            {

                Console.WriteLine(logFile);
                //clear all data
                readTo = false;
                readFrom = false;
                chngLog = false;
                linesFrom.Clear();
                linesToOld.Clear();
                linesToNew.Clear();
                linesLog.Clear();
                dataTo = null;
                dataFrom = null;
                dataFile = logFile.Replace("LOG", "HMI_NEW");//working graphic diagram file
                //if graphic diagram does not exist, get the next one

                if (!File.Exists(dataFile))
                    continue;
                
                directoryUPD = dataFile.Replace("\\HMI_NEW\\", "\\HMI_UPDATE\\");
                //if update file already exists delete it, if directory does not exists create it
                if (File.Exists(directoryUPD))
                    File.Delete(directoryUPD);
                if (!Directory.Exists(directoryUPD.Substring(0, directoryUPD.LastIndexOf("\\"))))
                    Directory.CreateDirectory(directoryUPD.Substring(0, directoryUPD.LastIndexOf("\\")));

                directoryLOGUPD = dataFile.Replace("\\HMI_NEW\\", "\\LOG_UPDATE\\");
                //if log update file already exists delete it, if directory does not exists create it
                if (File.Exists(directoryLOGUPD))
                    File.Delete(directoryLOGUPD);
                if (!Directory.Exists(directoryLOGUPD.Substring(0, directoryLOGUPD.LastIndexOf("\\"))))
                    Directory.CreateDirectory(directoryLOGUPD.Substring(0, directoryLOGUPD.LastIndexOf("\\")));
                //read all logged data and check if update points are used
                foreach (string line in File.ReadLines(@logFile, Encoding.GetEncoding(1252)))
                {
                    if (line.Contains("\"4"))
                        foreach (string point in pointsFrom)
                            if (line.Contains(point))
                                readTo = true;
                    //store all log file lines for later
                    linesLog.Enqueue(line);
                }
                //if update points are actually found, process the file
                if (readTo == true)
                {
                    Console.WriteLine(dataFile);
                    readTo = false;
                    //process graphic diagrams
                    foreach (string line in File.ReadLines(@dataFile, Encoding.GetEncoding(1252)))
                    {
                        actLine = line;
                        //copy logged data for replacement one set at a time
                        if (linesToOld.Count == 0)
                        {
                            procReplace = false;
                            {
                                Console.WriteLine(linesLogUPD.Count);
                                //sort all logged data
                                tempLine = (string)linesLog.Dequeue();
                                if (tempLine.Length != 0 || (!readFrom && !readTo))
                                {
                                    if (readTo)
                                    {
                                        //store lines to temporary storage
                                        tempQueue[modifications] = tempLine;
                                        foreach (string point in pointsFrom)
                                            if (tempQueue[modifications].Contains(point))
                                                chngLog = true;
                                        modifications++;
                                    }

                                    if (!readTo && readFrom && tempLine == "#to:")
                                    {
                                        readTo = true;
                                        readFrom = false;
                                        linesLogUPD.Enqueue(tempLine);
                                    }
                                    if (readFrom)
                                    {
                                        linesFrom.Enqueue(tempLine);
                                        linesLogUPD.Enqueue(tempLine);
                                    }

                                    if (!readFrom && tempLine == "#from:")
                                    {
                                        readFrom = true;
                                        linesLogUPD.Enqueue(tempLine);
                                    }
                                }
                                else
                                {
                                    if (linesFrom.Count > 0)
                                    {//replace all relevant points and make them IDs, not text
                                        for (int lp1=0;lp1<modifications;lp1++)
                                        {
                                            tempLine1 = tempQueue[lp1];
                                            if (chngLog)
                                            {
                                                //search through all points read from excel
                                                for (int lp2 = 0; lp2 < pointsFrom.Length; lp2++)
                                                    if (tempLine1.Contains(pointsFrom[lp2]))
                                                    {
                                                        tempLine1 = tempLine1.Replace('"' + pointsFrom[lp2] + '"', '/' + pointsTo[lp2]) + '/';
                                                        break;
                                                    }
                                            //when macro header line found, swap text argument number with point IDs number
                                                if (tempLine1.StartsWith("Macro") || tempLine1.StartsWith("MACRO"))
                                                {
                                                    tempLine1 = tempLine1.Replace("STATIC", "MINI");
                                                    linesToOld.Enqueue(tempQueue[lp1]);
                                                    tempLine1 = tempLine1.Replace(
                                                        dataTo[6] + " " + dataTo[7] + " " + dataTo[8] + " " + dataTo[9] + " " + dataTo[10] + " " + dataTo[11],
                                                        dataTo[7] + " " + dataTo[6] + " " + dataTo[8] + " " + dataTo[9] + " " + dataTo[10] + " " + dataTo[11]);
                                                    linesToNew.Enqueue(tempLine1);
                                                }
                                            }
                                            else
                                            {
                                                linesToOld.Enqueue(tempLine1);
                                                linesToNew.Enqueue(tempLine1);
                                            }
                                            modifications = 0;
                                            dataTo = tempLine1.Split(' ');
                                            linesLogUPD.Enqueue(tempLine1);
                                        }
                                        linesLogUPD.Enqueue(tempLine);
                                    }
                                    readFrom = false;
                                    break;
                             }
                         }
                      }
                      try
                      {
                        //split line read form src file
                        dataFrom = null;
                        dataFrom = actLine.Split(' ');
                        //ignore commented line
                        if (dataFrom[0].StartsWith("*") && actLine == (string)linesToOld.Peek()) { }
                        else
                        if (dataFrom.Length > 5)
                            if (procReplace || (dataFrom[2] == dataTo[2] && dataFrom[3] == dataTo[3] && dataFrom[4] == dataTo[4] && dataFrom[5] == dataTo[5]))//for macro replacemen
                            {
                                procReplace = true;
                                linesFrom.Dequeue();
                                actLine = (string)linesToNew.Dequeue();
                                linesToOld.Dequeue();
                            }
                          File.AppendAllText(directoryUPD, actLine + "\r\n", Encoding.GetEncoding(1252));
                        }
                    catch { }

                    }

                    // Console.WriteLine(dataFile.Replace("\\HMI_NEW\\", "\\LOG_UPDATE\\"));
                    while (linesLogUPD.Count > 0)
                        File.AppendAllText(directoryLOGUPD, linesLogUPD.Dequeue() + "\r\n", Encoding.GetEncoding(1252));
                }
            }
        }
    }
}



