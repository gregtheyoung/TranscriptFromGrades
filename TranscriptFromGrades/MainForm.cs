// Add a comment
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using Excel;
using Microsoft.Office.Interop.Word;

namespace TranscriptFromGrades
{
    public partial class MainForm : Form
    {
        private string rtfTemplateText;
        private string studentFullName;

        public MainForm()
        {
            InitializeComponent();

            rtfTemplateFileTextBox.Text = Properties.Settings.Default.TemplateFile;
        }

        private void openFileButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                excelFileTextBox.Text = openFileDialog.FileName;
        }

        private void templateFileButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                rtfTemplateFileTextBox.Text = openFileDialog.FileName;

        }

        private void generateTranscriptButton_Click(object sender, EventArgs e)
        {
            // Save off the name of the template file in the user settings so it is the same the next
            // time the user starts the app.
            Properties.Settings.Default.TemplateFile = rtfTemplateFileTextBox.Text;
            Properties.Settings.Default.Save();

            GenerateTranscript(excelFileTextBox.Text, rtfTemplateFileTextBox.Text);
        }

        private void GenerateTranscript(string excelFileName, string rtfTemplateFileName)
        {
            DataSet ds = GetDataSetFromExcelFile(excelFileName);
            rtfTemplateText = GetTranscriptRTFTemplateText(rtfTemplateFileName);

            // Get the name of the student for later use.
            studentFullName = ds.Tables["Info"].Select("FirstColumn = 'Name'")[0][1].ToString();

            ReplaceStudentInfo(ds);

            ReplaceAcademicRecord(ds);

            RemoveRemainingTags();

            CreateTranscript(Path.GetDirectoryName(excelFileName));
        }

        private void RemoveRemainingTags()
        {
            // Get rid of the rest of the tags that haven't been replaced - there isn't any data left to fill them.
            while (rtfTemplateText.IndexOf("[[") >= 0)
            {
                int startLocation = rtfTemplateText.IndexOf("[[");
                int endLocation = rtfTemplateText.IndexOf("]]");
                rtfTemplateText = rtfTemplateText.Remove(startLocation, endLocation - startLocation + 2);
            }
        }

        private void ReplaceAcademicRecord(DataSet ds)
        {
            // Get the one table from teh dataset we'll be using.
            System.Data.DataTable gradesTable = ds.Tables["Grades"];
            // Get rid of rows that have blanks...
            foreach (DataRow row in gradesTable.Rows)
            {
                if ((row["Subject"].ToString().Length == 0) ||
                    (row["Title"].ToString().Length == 0) ||
                    (row["Final"].ToString().Length == 0) ||
                    (row["Credit"].ToString().Length == 0))
                    row.Delete();
            }
            gradesTable.AcceptChanges();

            // Get a list of the school years to be used. Note that this should be from 1 to 4 rows,
            // of the format like "2012-2013".
            var distinctYears = (from row in gradesTable.AsEnumerable() select row.Field<string>("SchoolYear")).Distinct();

            // Use a DataView against the table so that we can sort the courses for presentation.
            DataView gradesView = gradesTable.AsDataView();
            gradesView.Sort = "SchoolYear ASC, Subject ASC, Title ASC";

            // This dictionary will contain a list of subjects, and for each one a SubjectCredit class, which contains
            // the summed credits for the year, the summed credits for all years, and the tag to find in the RTF text
            // for the subject.
            Dictionary<string, SubjectCredit> subjectCredits = InitializeSubjectCredits();
            double totalCredits = 0.0;
            double totalPoints = 0.0;

            bool hasCollege = false;
            bool hasTransfer = false;

            // Iterate through the list of distinct years...
            int sectionNumber = 1;
            bool first = true;
            foreach (string year in distinctYears)
            {
                if (!String.IsNullOrEmpty(year))
                {
                    rtfTemplateText = rtfTemplateText.Replace("[[year" + sectionNumber + "]]", year);

                    gradesView.RowFilter = "SchoolYear = '" + year + "'";

                    var rows = gradesView.ToTable().AsEnumerable();


                    // Obtain the RTF string that represents one row. We will do this by going from 
                    // a known tag in row 2 to the same known tag in row 3, and grabbing all
                    // the text inbetween. Then we will either delete it or repeat it.
                    int firstLocation = rtfTemplateText.IndexOf("[[year" + sectionNumber + "subject]]");
                    int secondLocation = rtfTemplateText.IndexOf("[[year" + sectionNumber + "subject]]", firstLocation + 1);
                    string clone = rtfTemplateText.Substring(firstLocation, secondLocation - firstLocation);

                    // If we have less than four rows, then we need to remove a row from the RTF template
                    if (rows.Count() < 4)
                        rtfTemplateText = rtfTemplateText.Remove(firstLocation, secondLocation - firstLocation);
                    else if (rows.Count() > 4)
                    {
                        // We need to add rows to the RTF template. We already have 4 in the template, so just
                        // need to add the difference.
                        for (int i = 0; i < rows.Count() - 4; i++)
                            rtfTemplateText = rtfTemplateText.Insert(secondLocation, clone);
                    }

                    double creditsThisYear = 0.0;
                    double pointsThisYear = 0.0;

                    // Now we iterate through each course in this year...
                    foreach (DataRow row in rows)
                    {
                        string patternPrefix = "[[year" + sectionNumber;
                        if (first) patternPrefix += "first";
                        if (rows.Last() == row) patternPrefix += "last";
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, patternPrefix + "subject]]", row["Subject"].ToString());
                        string title = row["Title"].ToString();
                        if ((row["CollegeYN"].ToString().Length > 0) && (row["CollegeYN"].ToString().ToLower()[0] == 'y'))
                        {
                            title += @" \super[C]";
                            hasCollege = true;
                        }
                        if ((row["TransferYN"].ToString().Length > 0) && (row["TransferYN"].ToString().ToLower()[0] == 'y'))
                        {
                            title += @" \super[T]";
                            hasTransfer = true;
                        }
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, patternPrefix + "title]]", title);
                        Double credit = new Double();
                        Double.TryParse(row["Credit"].ToString(), out credit);
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, patternPrefix + "credit]]", credit.ToString("f1"));
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, patternPrefix + "grade]]", row["Final"].ToString());

                        // Add up credits points for GPA computation.
                        switch (row["Final"].ToString().ToUpper()[0])
                        {
                            case 'A':
                                pointsThisYear += 4.0 * credit;
                                break;
                            case 'B':
                                pointsThisYear += 3.0 * credit;
                                break;
                            case 'C':
                                pointsThisYear += 2.0 * credit;
                                break;
                        }

                        subjectCredits[row["Subject"].ToString()].creditsThisYear += credit;
                        subjectCredits[row["Subject"].ToString()].totalCredits += credit;

                        creditsThisYear += credit;
                        totalCredits += credit;
                        first = false;
                    }

                    if (creditsThisYear > 0)
                    {
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, "[[year" + sectionNumber + "totalcredits]]", creditsThisYear.ToString("f1"));
                        Double gpaThisYear = pointsThisYear / creditsThisYear;
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, "[[year" + sectionNumber + "gpa]]", gpaThisYear.ToString("f2"));
                    }
                    else
                    {
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, "[[year" + sectionNumber + "totalcredits]]", "");
                        rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, "[[year" + sectionNumber + "gpa]]", "");
                    }

                    // Now replace the credit sums for each subject for this year...
                    foreach (string key in subjectCredits.Keys)
                    {
                        if (subjectCredits[key].creditsThisYear == 0.0)
                            rtfTemplateText = rtfTemplateText.Replace("[[year" + sectionNumber + subjectCredits[key].matchString + "]]", "");
                        else
                            rtfTemplateText = rtfTemplateText.Replace("[[year" + sectionNumber + subjectCredits[key].matchString + "]]", subjectCredits[key].creditsThisYear.ToString("f1"));
                        subjectCredits[key].creditsThisYear = 0.0;
                    }

                    totalPoints += pointsThisYear;
                    first = true;
                    sectionNumber++;
                }

            }

            // Replace the credit totals for each subject...
            foreach (string key in subjectCredits.Keys)
            {
                if (subjectCredits[key].totalCredits == 0.0)
                    rtfTemplateText = rtfTemplateText.Replace("[[total" + subjectCredits[key].matchString + "]]", "");
                else
                    rtfTemplateText = rtfTemplateText.Replace("[[total" + subjectCredits[key].matchString + "]]", subjectCredits[key].totalCredits.ToString("f1"));
            }

            // Notes if needed
            if (hasCollege || hasTransfer)
            {
                string notes = "";
                if (hasCollege) notes += @"[C] - denotes course was taken through a college\par";
                if (hasTransfer) notes += @"[T] - denotes course was transferred from another school";
                rtfTemplateText = ReplaceFirstPatternInstanceWithString(rtfTemplateText, "[[notes]]", notes);
            }

            rtfTemplateText = rtfTemplateText.Replace("[[totalcredits]]", totalCredits.ToString("f1"));
            double overallGPA = totalPoints / totalCredits;
            rtfTemplateText = rtfTemplateText.Replace("[[overallgpa]]", overallGPA.ToString("f2"));

        }

        private static Dictionary<string, SubjectCredit> InitializeSubjectCredits()
        {
            Dictionary<string, SubjectCredit> subjectCredits = new Dictionary<string, SubjectCredit>();
            subjectCredits.Add("Bible", new SubjectCredit { matchString = "bible", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("English", new SubjectCredit { matchString = "english", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Fine Arts", new SubjectCredit { matchString = "finearts", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Foreign Language", new SubjectCredit { matchString = "foreign", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Mathematics", new SubjectCredit { matchString = "math", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Other", new SubjectCredit { matchString = "other", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Physical Education", new SubjectCredit { matchString = "pe", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Science", new SubjectCredit { matchString = "science", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Service", new SubjectCredit { matchString = "service", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Social Studies", new SubjectCredit { matchString = "socialstudies", creditsThisYear = 0.0, totalCredits = 0.0 });
            subjectCredits.Add("Technology / Trade / Business", new SubjectCredit { matchString = "tech", creditsThisYear = 0.0, totalCredits = 0.0 });
            return subjectCredits;
        }

        private static string ReplaceFirstPatternInstanceWithString(string originalString, string pattern, string replacement)
        {
            string newString = originalString;

            int start = newString.IndexOf(pattern);
            if (start >= 0)
            {
                newString = newString.Remove(start, pattern.Length);
                newString = newString.Insert(start, replacement);
            }

            return newString;
        }

        private void ReplaceStudentInfo(DataSet ds)
        {
            DataRow row = ds.Tables["Info"].Select("FirstColumn = 'Name'")[0];
            SubstituteSingleString(row, @"[[fullname]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'DOB'")[0];
            SubstituteDate(row, @"[[birthdate]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'Parents'")[0];
            SubstituteSingleString(row, @"[[parentnames]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'Address1'")[0];
            SubstituteSingleString(row, @"[[addressline1]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'City'")[0];
            SubstituteSingleString(row, @"[[city]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'State'")[0];
            SubstituteSingleString(row, @"[[state]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'Zip'")[0];
            SubstituteSingleString(row, @"[[zip]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'Phone'")[0];
            SubstituteSingleString(row, @"[[phone]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'Email'")[0];
            SubstituteSingleString(row, @"[[email]]");

            row = ds.Tables["Info"].Select("FirstColumn = 'Grad Date'")[0];
            SubstituteDate(row, @"[[graddate]]");
            if (row[1].ToString() != "")
                rtfTemplateText = rtfTemplateText.Replace("[[diplomaearned]]", "Yes");
            else
                rtfTemplateText = rtfTemplateText.Replace("[[diplomaearned]]", "No");

        }

        private void SubstituteSingleString(DataRow row, string pattern)
        {
            string newString = row[1].ToString();
            rtfTemplateText = rtfTemplateText.Replace(pattern, newString);
        }

        private void SubstituteDate(DataRow row, string pattern)
        {
            string newString = row[1].ToString();

            DateTime dateTime = new DateTime();
            if (DateTime.TryParse(newString, out dateTime))
                newString = dateTime.ToShortDateString();
            else
                newString = "";

            rtfTemplateText = rtfTemplateText.Replace(pattern, newString);
        }

        private void SubstituteNumber(DataRow row, string pattern, int numberOfDigitsAfterDecimalPoint)
        {
            string newString = row[1].ToString();
            
            Double doubleValue = new Double();
            if (Double.TryParse(newString, out doubleValue))
                newString = doubleValue.ToString("f"+numberOfDigitsAfterDecimalPoint);
            else
                newString = "";

            rtfTemplateText = rtfTemplateText.Replace(pattern, newString);
        }

        /// <summary>
        /// Create the transcript files - both RTF and PDF.
        /// </summary>
        /// <param name="defaultDirectory"></param>
        private void CreateTranscript(string defaultDirectory)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = defaultDirectory;

            // Pattern for transcript file will be "FullName - Transcript - yyyy-mm-dd.rtf"
            saveFileDialog.FileName = studentFullName + " - Transcript " + DateTime.Now.ToString("o").Remove(10) + ".rtf";
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Pretty simple to write it out to RTF.
                File.WriteAllText(saveFileDialog.FileName, rtfTemplateText);

                CreatePDFFromRTF(saveFileDialog.FileName);
            }
        }

        private static void CreatePDFFromRTF(string rtfFileName)
        {
            // Use MS Word to open the RTF file and do a SaveAs to PDF.
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;
            Object oRTFFileName = (Object)rtfFileName;
            Document doc = word.Documents.Open(ref oRTFFileName);
            doc.Activate();

            object pdfFileName = rtfFileName.Replace(".rtf", ".pdf");
            object fileFormat = WdSaveFormat.wdFormatPDF;
            doc.SaveAs2(ref pdfFileName, ref fileFormat);

            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)doc).Close(ref saveChanges);
        }

        private static DataSet GetDataSetFromExcelFile(string fileName)
        {
            FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            excelReader.IsFirstRowAsColumnNames = false;
            DataSet ds = excelReader.AsDataSet();

            ds.Tables["Info"].Columns[0].ColumnName = "FirstColumn";
            ds.Tables["Grades"].Columns[0].ColumnName = "SchoolYear";
            ds.Tables["Grades"].Columns[1].ColumnName = "Subject";
            ds.Tables["Grades"].Columns[2].ColumnName = "Title";
            ds.Tables["Grades"].Columns[7].ColumnName = "Final";
            ds.Tables["Grades"].Columns[8].ColumnName = "Credit";
            ds.Tables["Grades"].Columns[17].ColumnName = "CollegeYN";
            ds.Tables["Grades"].Columns[18].ColumnName = "TransferYN";
            ds.Tables["Grades"].Rows[0].Delete();
            ds.AcceptChanges();

            stream.Close();
            excelReader.Close();

            return ds;
        }

        private static string GetTranscriptRTFTemplateText(string fileName)
        {
            StreamReader reader = File.OpenText(fileName);
            string allText = reader.ReadToEnd();

            reader.Close();

            return allText;
        }

    }

    class SubjectCredit
    {
        public string matchString;
        public double creditsThisYear;
        public double totalCredits;
    }
}
