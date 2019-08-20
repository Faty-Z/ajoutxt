using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;



using System.IO;
using System.Diagnostics;

namespace ajoutxt
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            this.Application.DocumentOpen += Application_DocumentOpen;
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += ThisAddIn_NewDocument;
        }
        
      public void Open_existingdoc()
        {
            string path = "C:\\Users\\PC\\Desktop\\test\\تأهيل ملعب رياضي للقرب بحي الزرهونية .docx";

            if (File.Exists(path))
                this.Application.Documents.Open(path);
            else
                Debug.Print("File not found @ {0}", path);
        }
        
        public void X1()
        {
            Word.Find findObject = Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "X1";
            findObject.Replacement.ClearFormatting();

           
            findObject.Replacement.Text = TextBox3.Text;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }
        public void X2()
        {
            Word.Find findObject = Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "X2";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = TextBox2.Text;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }
        public void X3()
        {
            Word.Find findObject = Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "X3";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = TextBox1.Text;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }
        public void X4()
        {
            Word.Find findObject = Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "X4";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = TextBox4.Text;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }
        public void X5()
        {
            Word.Find findObject = Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "X5";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = ListBox.Text;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

        private void ThisAddIn_NewDocument(Word.Document Doc)
        {
        }
        private void Application_DocumentOpen(Word.Document Doc)
        {
        }

        public void WorkWithDocument(Word.Document doc)
        {
            // Using InsertBefore method inserts text   
            doc.Content.InsertBefore("Text @ the          Start - ");
            // Using InsertAfter method inserts text   
            doc.Content.InsertAfter(" - Text @          the End");
        }

      
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
