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
using System.Text.RegularExpressions;
using System.Diagnostics;


namespace LeanenkaHolsted
{
    public partial class Form1 : Form
    {
        #region Function definition
        public static Regex RegexForVariables()
        {
            Regex VariablesMatch = new Regex(@"(?<=((Dim|Public|Protected|Friend|Protected|Friend|Private|Shared|Shadows|Static|ReadOnly|WithEvents){1,}\s))[\d\w_]{1,}",RegexOptions.IgnoreCase);
            return VariablesMatch;
        }
        public static List<string> SearchForPVariables(List<string> ListOfVariablesMatched,string Code)
        {
            List<string> ListOfPVariables = new List<string>();
            foreach (string VariableCandidateP in ListOfVariablesMatched)
            {
                string RexExp_P = @"[^\d\w_]"+VariableCandidateP+@"[^\d\w_][\d\D\w\W]{1,}";//non greeding repetition - any symbols until we match the name of the variable
                Regex RegForP_Variable = new Regex(RexExp_P,RegexOptions.IgnoreCase);
                Match First_P =  RegForP_Variable.Match(Code);
                if (!First_P.Success) { }
                else
                {   //Reusing match and regex
                    int IndexInCode = Code.IndexOf(First_P.Value);                    
                    string PartOfVBCodeUntillVariableisMatched = Code.Substring(0, IndexInCode);
                    RegForP_Variable = new Regex(@"[^\d\w_]Function[^\d\w_]|[^\d\w_]function[^\d\w_]", RegexOptions.IgnoreCase);
                    First_P = RegForP_Variable.Match(PartOfVBCodeUntillVariableisMatched);
                    if (First_P.Success)//P is not global - inside function
                    {

                    }
                    else
                    {
                        ListOfPVariables.Add(VariableCandidateP);
                    }
                }                    
            }
            return ListOfPVariables;
        }
        public static List<string> SearchForMVariables(List<string> ListOfVariablesMatched, string Code)// modified means that a global variable is changed or local variable is created and initialized
        {
            List<string> ListOfMVariables = new List<string>();
            foreach (string VariableCandidateM in ListOfVariablesMatched)
            {
                //declared - then initialized
                string RexExp_M = @"([^\d\w_]" + VariableCandidateM + @"\s{0,}=){1,1}?";//non greeding repetition - any symbols until we match the name of the variable
                Regex RegForM_Variable = new Regex(RexExp_M, RegexOptions.IgnoreCase);
                Match First_M = RegForM_Variable.Match(Code);
                if (First_M.Success) { ListOfMVariables.Add(VariableCandidateM); }//M class because of = sign
                else
                {
                    //declared + initialized 
                    RexExp_M = @"([^\d\w_]" + VariableCandidateM + @"[^\d\w_])([\d\w\s_]{1,})=";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                    RegForM_Variable = new Regex(RexExp_M, RegexOptions.IgnoreCase);
                    First_M = RegForM_Variable.Match(Code);
                    if (First_M.Success) { ListOfMVariables.Add(VariableCandidateM); }//M class because of = sign
                }
            }
            return ListOfMVariables;
        }
        public static List<string> SearchForCVariables(List<string> ListOfVariablesMatched, string Code)// modified means that a global variable is changed or local variable is created and initialized
        {
            List<string> ListOfCVariables = new List<string>();
            foreach (string VariableCandidateC in ListOfVariablesMatched)
            {
                //declared - then initialized
                string While_C = @"((?<=While\s{1,})"+ VariableCandidateC + @"){1,1}?";
                Regex RegForC_Variable = new Regex(While_C, RegexOptions.IgnoreCase);
                Match First_C = RegForC_Variable.Match(Code);
                if (First_C.Success) { ListOfCVariables.Add(VariableCandidateC); continue; }
                else
                {
                    //declared + initialized 
                    string Until_C = @"((?<=Until\s{1,})"+VariableCandidateC+@"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                    RegForC_Variable = new Regex(Until_C, RegexOptions.IgnoreCase);
                    First_C = RegForC_Variable.Match(Code);
                    if (First_C.Success) { ListOfCVariables.Add(VariableCandidateC); continue; }
                    else
                    {
                        string For_C = @"((?<=For\s{1,})" + VariableCandidateC + @"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                        RegForC_Variable = new Regex(For_C, RegexOptions.IgnoreCase);
                        First_C = RegForC_Variable.Match(Code);
                        if (First_C.Success) { ListOfCVariables.Add(VariableCandidateC); continue; }
                        else
                        {
                            string If_C = @"((?<=If\s{1,})" + VariableCandidateC + @"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                            RegForC_Variable = new Regex(If_C, RegexOptions.IgnoreCase);
                            First_C = RegForC_Variable.Match(Code);
                            if (First_C.Success) { ListOfCVariables.Add(VariableCandidateC); continue; }
                        }
                    }
                }
            }
            return ListOfCVariables;
        }

        #endregion
        #region Catching errors
        public static void CatchFunction(Exception ex)
        {
            StreamWriter Strw;
            FileInfo Errors = new FileInfo(Directory.GetCurrentDirectory() + @"\Errors.txt");
            if (!Errors.Exists)
                Strw = new StreamWriter(Errors.Create(), System.Text.Encoding.UTF8);
            else
                Strw = Errors.AppendText();
            Strw.WriteLine("\r\n*Errors logging at: {0}*", DateTime.Now);
            Strw.WriteLine("Message: " + ex.Message + "; source: " + ex.Source + "; data: " + ex.Data + "; stacktrace: " + ex.StackTrace);
            Strw.Flush();
            MessageBox.Show("An error has occured: " + ex.Message + ". For more information see the Errors.txt logs.");
            Application.Exit();
        }
        #endregion
        #region Events handler
        public Form1()
        {
            InitializeComponent();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            openFileDialog1.Filter = " VB files(*.vb, *.VB) | *.vb; *.VB | Text files(*.txt) | *.txt  | All files(*.*) | *.*";

            this.richTextBox1.Enabled = true;
            LinkLabel.Link link = new LinkLabel.Link();
	        link.LinkData = "https://msdn.microsoft.com/ru-ru/library/7ee5a7s1.aspx";//information about variables in VB 6
	        linkLabel1.Links.Add(link);

          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int size = 0;
            string file = string.Empty;
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(file);
                    size = text.Length;
                    richTextBox1.AppendText(text);
                }
                catch (Exception ex)
                {
                    CatchFunction(ex);
                }
            }
            label1.Text ="Size: "+ size+ " bytes;"; // <-- Shows file size in debugging mode.
            label2.Text = "Path: " + file + " ;"; // <-- For debugging use.
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            MessageBox.Show("Thank you for the program usage");
            Application.Exit();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            label1.Text = @"0 bytes;";
            label2.Text = @"' ';";
            for (int i = dataGridView1.Rows.Count-1; i > 0 ; --i)
                dataGridView1.Rows.RemoveAt(i-1);
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

         
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Regex RegExpOfVariables = RegexForVariables();
            string AnalysedCodeFromRichTextBox = richTextBox1.Text;
            var MatchedVariables = RegExpOfVariables.Matches(AnalysedCodeFromRichTextBox);
            List<string> ListOfVariablesMatched = new List<string>();
            for (int i = 0; i < MatchedVariables.Count; ++i)
                if(MatchedVariables[i].Value!="function" && MatchedVariables[i].Value != "Function")
                    ListOfVariablesMatched.Add( MatchedVariables[i].Value );

            List<string> P_Variables;
            List<string> M_Variables;
            List<string> C_Variables;
            List<string> T_Variables;
            P_Variables = SearchForPVariables(ListOfVariablesMatched, AnalysedCodeFromRichTextBox);
            M_Variables = SearchForMVariables(ListOfVariablesMatched, AnalysedCodeFromRichTextBox);
            C_Variables = SearchForCVariables(ListOfVariablesMatched, AnalysedCodeFromRichTextBox);

            List<string> ListOfP_M_C = new List<string>(P_Variables);
            ListOfP_M_C.AddRange(M_Variables);
            ListOfP_M_C.AddRange(C_Variables);
            ListOfP_M_C.Select(w => w).Distinct();
            T_Variables = ListOfVariablesMatched.Except(ListOfP_M_C.ToArray()).ToList();

            int QuantityOfRowsInTheTable = 0;
            if (P_Variables.Count > QuantityOfRowsInTheTable)
                QuantityOfRowsInTheTable = P_Variables.Count;
            
            if (M_Variables.Count > QuantityOfRowsInTheTable)
                QuantityOfRowsInTheTable = M_Variables.Count;
          
            if (C_Variables.Count > QuantityOfRowsInTheTable)
                QuantityOfRowsInTheTable = C_Variables.Count;
          
            if (T_Variables.Count > QuantityOfRowsInTheTable)
                QuantityOfRowsInTheTable = T_Variables.Count;

            for (int i = 0; i < QuantityOfRowsInTheTable + 1; ++i) // Extra one for computations
            {
                DataGridViewRow RowToBeAdded = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                dataGridView1.Rows.Add(RowToBeAdded);
            }
            dataGridView1.Rows[0].Cells[0].Value = P_Variables.Count;
            dataGridView1.Rows[0].Cells[1].Value = M_Variables.Count;
            dataGridView1.Rows[0].Cells[2].Value = C_Variables.Count;
            dataGridView1.Rows[0].Cells[3].Value = T_Variables.Count;
            dataGridView1.Rows[0].Cells[4].Value = P_Variables.Count + 2 *  M_Variables.Count + 3 * C_Variables.Count + 0.5 * T_Variables.Count;
            //addition of P class
            for (int i = 1; i <= P_Variables.Count ; ++i)
                dataGridView1.Rows[i].Cells[0].Value = P_Variables[i-1];
            //addition of M class
            for (int i = 1; i <= M_Variables.Count; ++i)
                dataGridView1.Rows[i].Cells[1].Value = M_Variables[i - 1];
            //addition of C class
            for (int i = 1; i <= C_Variables.Count; ++i)
                dataGridView1.Rows[i].Cells[2].Value = C_Variables[i - 1];
            //addition of T class
            for (int i = 1; i <= T_Variables.Count; ++i)
                dataGridView1.Rows[i].Cells[3].Value = T_Variables[i - 1];
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(e.Link.LinkData as string);
        }
        #endregion
    }
}
