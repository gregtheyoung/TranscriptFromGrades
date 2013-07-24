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

namespace TranscriptFromGrades
{
    public partial class MainForm : Form
    {
        private string template;
        private string fullName;

        public MainForm()
        {
            InitializeComponent();

            templateFileTextBox.Text = Properties.Settings.Default.TemplateFile;
        }

        private void openFileButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                excelFileTextBox.Text = openFileDialog.FileName;
            }
        }

        private void templateFileButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                templateFileTextBox.Text = openFileDialog.FileName;
            }

        }

        private void generateTranscriptButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.TemplateFile = templateFileTextBox.Text;
            Properties.Settings.Default.Save();

            GenerateTranscript(excelFileTextBox.Text, templateFileTextBox.Text);
        }

        private void GenerateTranscript(string excelFileName, string templateFileName)
        {
            DataSet ds = GetDataSetFromExcelFile(excelFileName);
            template = GetTranscriptTemplate(templateFileName);

            fullName = ds.Tables["Info"].Select("FirstColumn = 'Name'")[0][1].ToString();

            ReplaceStudentInfo(ds);

            ReplaceAcademicRecord(ds);

            CreateTranscript(Path.GetDirectoryName(excelFileName));
        }

        private void ReplaceAcademicRecord(DataSet ds)
        {
            DataTable gradesTable = ds.Tables["Grades"];

            DataView gradesView = gradesTable.AsDataView();

            gradesView.Sort = "SchoolYear ASC, Subject ASC, Title ASC";

            var distinctYears = (from row in gradesTable.AsEnumerable() select row.Field<string>("SchoolYear")).Distinct();

            Dictionary<string, SubjectCredit> subjectCredits = new Dictionary<string, SubjectCredit>();
            subjectCredits.Add("Bible", new SubjectCredit {matchString="bible", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("English", new SubjectCredit {matchString="english", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Fine Arts", new SubjectCredit {matchString="finearts", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Foreign Language", new SubjectCredit {matchString="foreign", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Mathematics", new SubjectCredit {matchString="math", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Other", new SubjectCredit {matchString="other", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Physical Education", new SubjectCredit {matchString="pe", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Science", new SubjectCredit {matchString="science", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Service", new SubjectCredit {matchString="service", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Social Studies", new SubjectCredit {matchString="socialstudies", creditsThisYear=0.0, totalCredits=0.0});
            subjectCredits.Add("Technology / Trade / Business", new SubjectCredit { matchString = "tech", creditsThisYear = 0.0, totalCredits = 0.0 });

            int sectionNumber = 1;
            bool first = true;
            foreach (string year in distinctYears)
            {
                if (!String.IsNullOrEmpty(year))
                {
                    template = template.Replace("[[year" + sectionNumber + "]]", year);

                    gradesView.RowFilter = "SchoolYear = '" + year + "'";

                    var rows = gradesView.ToTable().AsEnumerable();

                    // Obtain the RTF string that represents one row. We will do this by going from 
                    // a known tag in row 2 to the same known tag in row 3, and grabbing all
                    // the text inbetween.
                    int firstLocation = template.IndexOf("[[year" + sectionNumber + "subject]]");
                    int secondLocation = template.IndexOf("[[year" + sectionNumber + "subject]]", firstLocation + 1);
                    string clone = template.Substring(firstLocation, secondLocation - firstLocation);

                    // If we have less than four rows, then we need to remove a row from the RTF template
                    if (rows.Count() < 4)
                        template = template.Remove(firstLocation, secondLocation - firstLocation);
                    else if (rows.Count() > 4)
                    {
                        // We need to add rows to the RTF template
                        for (int i = 0; i < rows.Count() - 4; i++)
                            template = template.Insert(secondLocation, clone);
                    }

                    double creditsThisYear = 0.0;
                    double pointsThisYear = 0.0;

                    foreach (DataRow row in rows)
                    {
                        string patternPrefix = "[[year" + sectionNumber;
                        if (first) patternPrefix += "first";
                        if (rows.Last() == row) patternPrefix += "last";
                        template = ReplaceFirstPatternInstanceWithString(template, patternPrefix + "subject]]", row["Subject"].ToString());
                        template = ReplaceFirstPatternInstanceWithString(template, patternPrefix + "title]]", row["Title"].ToString());
                        Double credit = new Double();
                        Double.TryParse(row["Credit"].ToString(), out credit);
                        template = ReplaceFirstPatternInstanceWithString(template, patternPrefix + "credit]]", credit.ToString("f1"));
                        template = ReplaceFirstPatternInstanceWithString(template, patternPrefix + "grade]]", row["Final"].ToString());
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
                        first = false;
                    }

                    template = ReplaceFirstPatternInstanceWithString(template, "[[year" + sectionNumber + "totalcredits]]", creditsThisYear.ToString("f1"));
                    Double gpaThisYear = pointsThisYear / creditsThisYear;
                    template = ReplaceFirstPatternInstanceWithString(template, "[[year" + sectionNumber + "gpa]]", gpaThisYear.ToString("f2"));

                    foreach (string key in subjectCredits.Keys)
                    {
                        if (subjectCredits[key].creditsThisYear == 0.0)
                            template = template.Replace("[[year" + sectionNumber + subjectCredits[key].matchString + "]]", "");
                        else
                            template = template.Replace("[[year" + sectionNumber + subjectCredits[key].matchString + "]]", subjectCredits[key].creditsThisYear.ToString("f1"));
                        subjectCredits[key].creditsThisYear = 0.0;
                    }

                    first = true;
                    sectionNumber++;
                }

            }
            foreach (string key in subjectCredits.Keys)
            {
                if (subjectCredits[key].totalCredits == 0.0)
                    template = template.Replace("[[total" + subjectCredits[key].matchString + "]]", "");
                else
                    template = template.Replace("[[total" + subjectCredits[key].matchString + "]]", subjectCredits[key].totalCredits.ToString("f1"));
            }

            // Get rid of the rest of the tags that haven't been replaced.
            while (template.IndexOf("[[") >= 0)
            {
                int startLocation = template.IndexOf("[[");
                int endLocation = template.IndexOf("]]");
                template = template.Remove(startLocation, endLocation - startLocation + 2);
            }
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
                template = template.Replace("[[diplomaearned]]", "Yes");
            else
                template = template.Replace("[[diplomaearned]]", "No");

            row = ds.Tables["Info"].Select("FirstColumn = 'Credits'")[0];
            SubstituteNumber(row, @"[[totalcredits]]", 1);

            row = ds.Tables["Info"].Select("FirstColumn = 'GPA'")[0];
            SubstituteNumber(row, @"[[overallgpa]]", 2);

        }

        private void SubstituteSingleString(DataRow row, string pattern)
        {
            string newString = row[1].ToString();
            template = template.Replace(pattern, newString);
        }

        private void SubstituteDate(DataRow row, string pattern)
        {
            string newString = row[1].ToString();

            DateTime dateTime = new DateTime();
            if (DateTime.TryParse(newString, out dateTime))
                newString = dateTime.ToShortDateString();
            else
                newString = "";

            template = template.Replace(pattern, newString);
        }

        private void SubstituteNumber(DataRow row, string pattern, int numberOfDigitsAfterDecimalPoint)
        {
            string newString = row[1].ToString();
            
            Double doubleValue = new Double();
            if (Double.TryParse(newString, out doubleValue))
                newString = doubleValue.ToString("f"+numberOfDigitsAfterDecimalPoint);
            else
                newString = "";

            template = template.Replace(pattern, newString);
        }

        private void CreateTranscript(string defaultDirectory)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = defaultDirectory;
            saveFileDialog.FileName = fullName + " - Transcript " + DateTime.Now.ToString("o").Remove(10) + ".rtf";
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                File.WriteAllText(saveFileDialog.FileName, template);
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
            ds.Tables["Grades"].Rows[0].Delete();
            ds.AcceptChanges();

            stream.Close();
            excelReader.Close();

            return ds;
        }

        private static string GetTranscriptTemplate(string fileName)
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
