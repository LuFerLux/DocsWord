using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
namespace AppAccounting
{
    class WordMerge
    {

        private const string defaultWordDocumentTemplate = @"Normal.dot";
        public static void Merge(string[] filesToMerge, string outputFilename, bool insertPageBreaks)
        {
            Merge(filesToMerge, outputFilename, insertPageBreaks, defaultWordDocumentTemplate);
        }
        /// <summary>
        /// A function that merges Microsoft Word Documents that uses a template specified by the user
        /// </summary>
        /// <param name="filesToMerge">An array of files that we want to merge</param>
        /// <param name="outputFilename">The filename of the merged document</param>
        /// <param name="insertPageBreaks">Set to true if you want to have page breaks inserted after each document</param>
        /// <param name="documentTemplate">The word document you want to use to serve as the template</param>
        public static void Merge(string[] filesToMerge, string outputFilename, bool insertPageBreaks, string documentTemplate)
        {
            object defaultTemplate = documentTemplate;
            object missing = System.Type.Missing;
            object pageBreak = Word.WdBreakType.wdPageBreak;
            object outputFile = outputFilename;

            // Create  a new Word application
            Word._Application wordApplication = new Word.Application();

            try
            {
                // Create a new file based on our template
                Word._Document wordDocument = wordApplication.Documents.Add(
                                              ref defaultTemplate
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                // Make a Word selection object.
                Word.Selection selection = wordApplication.Selection;

                // Loop thru each of the Word documents
                foreach (string file in filesToMerge)
                {
                    // Insert the files to our template
                    selection.InsertFile(
                                                file
                                            , ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                    //Do we want page breaks added after each documents?
                    if (insertPageBreaks)
                    {
                        selection.InsertBreak(ref pageBreak);
                    }
                }

                // Save the document to it's output file.
                wordDocument.SaveAs(
                                ref outputFile
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing);

                // Clean up!
                wordDocument = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }
        }
    }
}
