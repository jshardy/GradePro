using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using word = Microsoft.Office.Interop.Word;

namespace GradePro
{
    public partial class frmMain : Form
    {
        //release version can use default app folder - remove path
        const string coursesFile = @"C:\Users\jshar\Desktop\Grading\Courses.txt";
        const string professorFile = @"C:\Users\jshar\Desktop\Grading\Professors.txt";
        const string settingsFile = @"C:\Users\jshar\Desktop\Grading\GradingPro.ini";
        const string homeworkFile = @"C:\Users\jshar\Desktop\Grading\Homework.txt";
        const string magicspace = "_____ ";
        const string magicx = "__✓__ ";
        public frmMain()
        {
            InitializeComponent();
            //Winforms is jacked up.. resize ourselves.
            Width = 755;
            Height = 760;

            StreamReader sr;
            //I don't feel like writing an INI class right now so multiple files it is.
            if(File.Exists(coursesFile))
            {
                sr = new StreamReader(coursesFile);

                while(!sr.EndOfStream)
                    cmboClass.Items.Add(sr.ReadLine());

                sr.Close();
            }

            if(File.Exists(professorFile))
            {
                sr = new StreamReader(professorFile);

                while(!sr.EndOfStream)
                    cmboProfessor.Items.Add(sr.ReadLine());

                sr.Close();
            }


            if(File.Exists(homeworkFile))
            {
                sr = new StreamReader(homeworkFile);

                while(!sr.EndOfStream)
                    cmboHomework.Items.Add(sr.ReadLine());

                sr.Close();
            }

            if(File.Exists(settingsFile))
            {
                sr = new StreamReader(settingsFile);
                txtPath.Text = sr.ReadLine();
                sr.Close();
            }

            cmboClass.SelectedIndex = 0;
            cmboProfessor.SelectedIndex = 0;
            cmboHomework.SelectedIndex = 0;
        }

        private void btnPath_Click(object sender, EventArgs e)
        {
            using(var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if(result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    txtPath.Text = fbd.SelectedPath;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            bool failed = false;

            btnSave.Enabled = false;
            btnSave.Text = "Working";
            //Everything has to be an object(COM) to pass to word.
            object strFileName = txtPath.Text + "\\" + cmboHomework.Text + " - " + cmboStudentName.Text + ".doc";
            try
            {
                File.Copy(txtPath.Text + "\\GradingTemplate.doc", (string)strFileName, true);
            }
            catch
            {
                failed = true;
                MessageBox.Show("Error can't access " + strFileName);
            }

            if(failed == false)
            {
                word.Application wordApp = new word.Application();
                wordApp.Visible = chkMSWord.Checked;
                wordApp.WindowState = word.WdWindowState.wdWindowStateNormal;

                object missing = System.Reflection.Missing.Value;
                object readOnly = false;
                object isVisible = true;

                word.Document doc = wordApp.Documents.Open(ref strFileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible);
                // Activate the document so it shows up in front  
                doc.Activate();


                SearchReplace("#1", cmboClass.Text, wordApp);
                SearchReplace("#2", cmboProfessor.Text, wordApp);
                SearchReplace("#3", cmboStudentName.Text, wordApp);
                SearchReplace("#4", cmboHomework.Text, wordApp);

                //This magic lets me use the checkbox name as the ms word field name
                //You can just add a checkbox to the form and a spot on the template
                //and it will automatically detect it with no code changes.
                foreach(var chkBox in Utility.GetAllChildren(this).OfType<CheckBox>())
                {
                    if(chkBox.Checked && chkBox != chkMSWord)
                    {
                        string newText = chkBox.Text;
                        if(chkBox.Text == "Late submission")
                            newText += " " + txtDaysLate.Text + " days";

                        SearchReplace(magicspace + chkBox.Text, magicx + newText, wordApp);
                    }
                }

                //insert the comments at the end.

                wordApp.Selection.EndKey(word.WdUnits.wdStory, ref missing);
                wordApp.Selection.TypeText("\r\n\r\nComments: " + txtComments.Text);

                doc.Save();

                //leave it open for inspection if chosen
                if(!chkMSWord.Checked)
                {
                    doc.Close();
                    wordApp.Quit();
                }
            }
            btnSave.Enabled = true;
            btnSave.Text = "Save";
            btnClear_Click(this, null);
        }

        private void SearchReplace(string SearchFor, string ReplaceWith, word.Application wordApp)
        {
            word.Find findObj = wordApp.Selection.Find;
            findObj.ClearFormatting();
            findObj.Text = SearchFor;
            findObj.Replacement.ClearFormatting();
            findObj.Replacement.Text = ReplaceWith;

            object missing = System.Reflection.Missing.Value;

            object replace = word.WdReplace.wdReplaceOne;
            findObj.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replace, ref missing, ref missing, ref missing, ref missing);

            //same thing as clicking away from a selection
            wordApp.Selection.Collapse();

            //put the cursor at start of document or find won't work.
            Object toWhat = word.WdGoToItem.wdGoToLine;
            Object toWhich = word.WdGoToDirection.wdGoToFirst;
            wordApp.Selection.GoTo(toWhat, toWhich, ref missing, ref missing);
        }

        //Only allow numbers to be entered in day.
        private void txtDaysLate_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
                e.Handled = true;
        }

        //Clear all the checkboxes.
        private void btnClear_Click(object sender, EventArgs e)
        {
            txtComments.Text = txtDaysLate.Text = cmboStudentName.Text = "";
            foreach(var chkBox in Utility.GetAllChildren(this).OfType<CheckBox>())
            {
                if(chkBox.Checked && chkBox != chkMSWord)
                    chkBox.Checked = false;
            }
        }
    }
}
