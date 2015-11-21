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
        public void CalculateAllVariables(string FunctionOrSub,List<string> P, List<string> M, List<string> C, List<string> T)//for a single function
        {
            Regex RegexAllVariables = GetAllVariables();            
            var MatchedVariables = RegexAllVariables.Matches(FunctionOrSub);
            List<string> ListAllVariables = new List<string>();
            for (int i = 0; i < MatchedVariables.Count; ++i)
                if (MatchedVariables[i].Value != "function" && MatchedVariables[i].Value != "Function" && MatchedVariables[i].Value != "Sub" && MatchedVariables[i].Value != "sub")
                    ListAllVariables.Add(MatchedVariables[i].Value);
            Regex RegexFunctionParameters = new Regex(@"(?<=(ByVal|ByRef)\s{1,}?)[\d\w_]{1,}");
            var Matched_P = RegexFunctionParameters.Matches(FunctionOrSub);
            for (int i = 0; i < Matched_P.Count; ++i)
                if (Matched_P[i].Value != "function" && Matched_P[i].Value != "Function" && Matched_P[i].Value != "Sub" && Matched_P[i].Value != "sub")
                {
                    ListAllVariables.Add(Matched_P[i].Value);
                    P.Add(Matched_P[i].Value);
                }
            SearchMVariables(ListAllVariables, FunctionOrSub, M);
            SearchCVariables(ListAllVariables, FunctionOrSub, C);
            List<string> P_M_C_InModule = new List<string>();
            P_M_C_InModule.AddRange(P);
            P_M_C_InModule.AddRange(M);
            P_M_C_InModule.AddRange(C);
            T.AddRange(ListAllVariables.Except(P_M_C_InModule));
        }
        public static void CalculateRest(string Rest, List<string> P, List<string> M, List<string> C, List<string> T)
        {
            Regex RegexAllVariables = GetAllVariables();
            var MatchedVariables = RegexAllVariables.Matches(Rest);
            List<string> ListAllVariables = new List<string>();
            for (int i = 0; i < MatchedVariables.Count; ++i)
                if (MatchedVariables[i].Value != "function" && MatchedVariables[i].Value != "Function" && MatchedVariables[i].Value != "Sub" && MatchedVariables[i].Value != "sub")
                {
                    ListAllVariables.Add(MatchedVariables[i].Value);
                    P.Add(MatchedVariables[i].Value);
                }
            SearchMVariables(ListAllVariables, Rest, M);           
           
        }
        public static Regex GetAllVariables()
        {
            Regex VariablesMatch = new Regex(@"(?<=((Dim|Public|Protected|Friend|Protected|Friend|Private|Shared|Shadows|Static|ReadOnly|WithEvents){1,}\s))[\d\w_]{1,}",RegexOptions.IgnoreCase);
            return VariablesMatch;
        }
        //public static List<string> Add_P_Variables(List<string> ListVariables, string VBCodeCopy, List<string> ListGlobalVariables)
        //{
        //    VBCodeCopy = Regex.Replace(VBCodeCopy, @"(?<=((Dim|Public|Protected|Friend|Protected|Friend|Private|Shared|Shadows|Static|ReadOnly|WithEvents){1,}\s))[\d\w_]{1,}", @"", RegexOptions.IgnoreCase);
        //    foreach (string Variable in ListVariables)
        //    {
        //        Regex Regex_P_Candidates = new Regex(@"(?<=[^\d\w_])" + Variable + @"(?=[^\d\w_])", RegexOptions.IgnoreCase);
        //        MatchCollection Match_P_Candidates = Regex_P_Candidates.Matches(VBCodeCopy);
        //        if (Match_P_Candidates.Count != 0)
        //            ListGlobalVariables.Add(Variable);
        //    }

        //    return ListGlobalVariables;
        //}

        //public static List<string> SearchGlobalVariables(List<string> ListVariablesMatched,string VBCode)//there is a suggestion that all the global variables might be before all the function declarations
        //{
        //    List<string> ListPVariables = new List<string>();
        //    foreach (string VariableCandidateP in ListVariablesMatched)
        //    {
        //        string RexExp_P = @"[^\d\w_]"+VariableCandidateP+@"[^\d\w_][\d\D\w\W]{1,}";//non greeding repetition - any symbols until we match the name of the variable
        //        Regex RegForP_Variable = new Regex(RexExp_P,RegexOptions.IgnoreCase);
        //        Match First_P =  RegForP_Variable.Match(VBCode);
        //        if (!First_P.Success) { }
        //        else
        //        {   //Reusing match and regex
        //            int IndexInCode = VBCode.IndexOf(First_P.Value);                    
        //            string PartOfVBCodeUntillVariableisMatched = VBCode.Substring(0, IndexInCode);
        //            RegForP_Variable = new Regex(@"[^\d\w_](Function|Sub)[^\d\w_]|[^\d\w_]function[^\d\w_]", RegexOptions.IgnoreCase);
        //            First_P = RegForP_Variable.Match(PartOfVBCodeUntillVariableisMatched);
        //            if (First_P.Success)//P is not global - inside function
        //            {

        //            }
        //            else
        //            {
        //                ListPVariables.Add(VariableCandidateP);
        //            }
        //        }                    
        //    }
        //    return ListPVariables;
        //}
        public static void SearchMVariables(List<string> ListVariablesMatched, string VBCode, List<string> M)// modified means that a global variable is changed or local variable is created and initialized
        {
           
            foreach (string VariableCandidateM in ListVariablesMatched)
            {
                //declared - then initialized
                string RexExp_M = @"([^\d\w_]" + VariableCandidateM + @".{0,}=){1,1}?";//non greeding repetition - any symbols until we match the name of the variable
                Regex RegForM_Variable = new Regex(RexExp_M, RegexOptions.IgnoreCase);
                Match First_M = RegForM_Variable.Match(VBCode);
                if (First_M.Success) { M.Add(VariableCandidateM); }//M class because of = sign
                else
                {
                    //declared + initialized 
                    RexExp_M = @"([^\d\w_]" + VariableCandidateM + @"[^\d\w_])([\d\w\s_]{1,})=";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                    RegForM_Variable = new Regex(RexExp_M, RegexOptions.IgnoreCase);
                    First_M = RegForM_Variable.Match(VBCode);
                    if (First_M.Success) { M.Add(VariableCandidateM); }//M class because of = sign
                }
            }
            return;
        }
        public static void SearchCVariables(List<string> ListVariablesMatched, string VBCode, List<string> C)// modified means that a global variable is changed or local variable is created and initialized
        {
            
            foreach (string VariableCandidateC in ListVariablesMatched)
            {
                //declared - then initialized
                string While_C = @"((?<=While\s{1,})"+ VariableCandidateC + @"){1,1}?";
                Regex RegForC_Variable = new Regex(While_C, RegexOptions.IgnoreCase);
                Match First_C = RegForC_Variable.Match(VBCode);
                if (First_C.Success) { C.Add(VariableCandidateC); continue; }
                else
                {
                    //declared + initialized 
                    string Until_C = @"((?<=Until\s{1,})"+VariableCandidateC+@"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                    RegForC_Variable = new Regex(Until_C, RegexOptions.IgnoreCase);
                    First_C = RegForC_Variable.Match(VBCode);
                    if (First_C.Success) { C.Add(VariableCandidateC); continue; }
                    else
                    {
                        string For_C = @"((?<=For\s{1,})" + VariableCandidateC + @"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                        RegForC_Variable = new Regex(For_C, RegexOptions.IgnoreCase);
                        First_C = RegForC_Variable.Match(VBCode);
                        if (First_C.Success) { C.Add(VariableCandidateC); continue; }
                        else
                        {
                            string If_C = @"((?<=If\s{1,})" + VariableCandidateC + @"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                            RegForC_Variable = new Regex(If_C, RegexOptions.IgnoreCase);
                            First_C = RegForC_Variable.Match(VBCode);
                            if (First_C.Success) { C.Add(VariableCandidateC); continue; }
                            else
                            {
                                string Each_C = @"((?<=For\s{1,}Each[\d\w_\s]{1,}In\s{1,})" + VariableCandidateC + @"){1,1}?";//(?<=\d\w)=){1,1}?non greeding repetition - any symbols until we match the name of the variable
                                RegForC_Variable = new Regex(Each_C, RegexOptions.IgnoreCase);
                                First_C = RegForC_Variable.Match(VBCode);
                                if (First_C.Success) { C.Add(VariableCandidateC); continue; }
                            }
                        }
                    }
                }
            }
            return;
        }

        #endregion
        #region Catching errors
        public static void CatchFunction(Exception ex)
        {
            StreamWriter WriteErrors;
            FileInfo Errors = new FileInfo(Directory.GetCurrentDirectory() + @"\Errors.txt");
            if (!Errors.Exists)
                WriteErrors = new StreamWriter(Errors.Create(), System.Text.Encoding.UTF8);
            else
                WriteErrors = Errors.AppendText();
            WriteErrors.WriteLine("\r\n*Errors logging at: {0}*", DateTime.Now);
            WriteErrors.WriteLine("Message: " + ex.Message + "; source: " + ex.Source + "; data: " + ex.Data + "; stacktrace: " + ex.StackTrace);
            WriteErrors.Flush();
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
	        link.LinkData = "http://www.academia.edu/5563523/Software_Complexity_Metrics_A_Survey";//information about variables in VB 6
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
            richTextBox2.Text = "";
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

         
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = dataGridView1.Rows.Count - 1; i > 0; --i)
                dataGridView1.Rows.RemoveAt(i - 1);
            richTextBox2.Text = "";
            string VBCodeWithCommentsAndStrings = richTextBox1.Text;
            Regex RegexComments = new Regex(@"('.{0,}\n){1,1}?");
            var AllCommentsMatch = RegexComments.Matches(VBCodeWithCommentsAndStrings);
            foreach(Match Comment in AllCommentsMatch)
            {
                VBCodeWithCommentsAndStrings = VBCodeWithCommentsAndStrings.Replace(Comment.Value.ToString(), @"");
            }

            Regex RegexAllStrings = new Regex("(\"(.*?)\"){1,1}?",RegexOptions.IgnoreCase);
            var AllStringsMatch = RegexAllStrings.Matches(VBCodeWithCommentsAndStrings);
            foreach (Match aString in AllStringsMatch)
            {
                VBCodeWithCommentsAndStrings = VBCodeWithCommentsAndStrings.Replace(aString.Value.ToString(), @"");
            }

            string VBCode = VBCodeWithCommentsAndStrings;//without comments and strings            
            richTextBox2.AppendText("//Without comments and strings//"); richTextBox2.AppendText("\n");
            richTextBox2.AppendText(VBCode);
            Regex RegexFunctionsSub = new Regex(@"(Sub[\d\D\w\W]{1,}?End\s{1,}?Sub)|(Function[\d\D\w\W]{1,}?End\s{1,}?Function)", RegexOptions.IgnoreCase);//|     (  ((?<=[^\d\w_])Function([\w\W\d\D.\s\n]){1,}?End\s{1,}Function){1,1}?  )  
            MatchCollection MatchFunctionsAndSubs = RegexFunctionsSub.Matches(VBCode);
            //all the functions calculated

            List<string> P_Variables = new List<string>();
            List<string> M_Variables = new List<string>();
            List<string> C_Variables = new List<string>();
            List<string> T_Variables = new List<string>();
            string VBCodeCopy = VBCode;

            for (int i = 0; i < MatchFunctionsAndSubs.Count; ++i)//for each module add variables
            {
                CalculateAllVariables(MatchFunctionsAndSubs[i].Value, P_Variables, M_Variables, C_Variables, T_Variables);
                VBCodeCopy = VBCodeCopy.Replace(MatchFunctionsAndSubs[i].Value, "");
                richTextBox2.AppendText("\n"); richTextBox2.AppendText(@"//Function\Sub//"); richTextBox2.AppendText("\n");
                richTextBox2.AppendText(MatchFunctionsAndSubs[i].Value);
                
            }
            richTextBox2.AppendText("\n"); richTextBox2.AppendText("//Rest part//"); richTextBox2.AppendText("\n");
            richTextBox2.AppendText(VBCodeCopy);

            CalculateRest(VBCodeCopy, P_Variables, M_Variables, C_Variables, T_Variables);



            //Regex RegexAllVariables = GetAllVariables();
            //var MatchedVariables = RegexAllVariables.Matches(VBCode);
            //List<string> ListAllVariables = new List<string>();
            //for (int i = 0; i < MatchedVariables.Count; ++i)
            //    if(MatchedVariables[i].Value!="function" && MatchedVariables[i].Value != "Function" && MatchedVariables[i].Value != "Sub" && MatchedVariables[i].Value != "sub")
            //        ListAllVariables.Add( MatchedVariables[i].Value );

            ////P_Variables = SearchPVariables(ListAllVariables, VBCode);
            //List<string> ListGlobalVariables = SearchGlobalVariables(ListAllVariables, VBCodeCopy);
            //M_Variables = SearchMVariables(ListAllVariables, VBCode);
            //C_Variables = SearchCVariables(ListAllVariables, VBCode);
            //P_Variables = Add_P_Variables(ListAllVariables, VBCodeCopy, ListGlobalVariables);
            //P_Variables = new List<string>( P_Variables.Select(w => w).Distinct());
            //List<string> ListP_M_C = new List<string>(P_Variables);
            //ListP_M_C.AddRange(M_Variables);
            //ListP_M_C.AddRange(C_Variables);
            //ListP_M_C = new List<string>( ListP_M_C.Select(w => w).Distinct());
            //T_Variables = ListAllVariables.Except(ListP_M_C.ToArray()).ToList();

            int QuantityRowsTable = 0;
            if (P_Variables.Count > QuantityRowsTable)
                QuantityRowsTable = P_Variables.Count;
            
            if (M_Variables.Count > QuantityRowsTable)
                QuantityRowsTable = M_Variables.Count;
          
            if (C_Variables.Count > QuantityRowsTable)
                QuantityRowsTable = C_Variables.Count;
          
            if (T_Variables.Count > QuantityRowsTable)
                QuantityRowsTable = T_Variables.Count;

            for (int i = 0; i < QuantityRowsTable + 1; ++i) // Extra one for computations
            {
                DataGridViewRow RowAdded = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                dataGridView1.Rows.Add(RowAdded);
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
