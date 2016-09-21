using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
namespace PlantillaWord
{
    class WordDocument
    {
        private Word.Application wordApp;
        private Word.Document aDoc;
        private object missing = Missing.Value;
        private object filename;
        public WordDocument(string docPath, string archivo, string temporal)
        {
            try{
                temporal = "\\"+temporal;
                File.Copy(docPath+archivo, docPath+temporal,true);
                docPath=docPath+temporal;
                wordApp = new Word.Application();
                filename = Path.Combine(docPath);
                if (File.Exists((string)filename)){
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;
                    aDoc = wordApp.Documents.Open(ref filename, ref missing,
                    ref readOnly, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref isVisible, ref missing, ref missing,
                    ref missing, ref missing);
                    aDoc.Activate();
                }else {
                    MessageBox.Show("El archivo no existe.", "Sin archivo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) {
                MessageBox.Show("Error durante el proceso. Descripcion: " + ex.Message, "Error Interno", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  
            }
        }

        public void SaveDocument(){
            try{
                aDoc.Save();
            }catch (Exception ex){
                MessageBox.Show("Error durante el proceso. Descripcion: " + ex.Message, "Error Interno", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CloseDocument()
        {
            object saveChanges = Word.WdSaveOptions.wdSaveChanges;
            wordApp.Quit(ref saveChanges,ref missing,ref missing);
        }

        public void FindAndReplace(object findText, object replaceText)
        {
            this.findAndReplace(this.wordApp, findText, replaceText);
        }

        public void findAndReplace(Word.Application wordApp, object findText, object replaceText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            if (wordApp!= null)
            {
                wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms,
                ref forward, ref wrap, ref format, ref replaceText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
        }
    }
}
