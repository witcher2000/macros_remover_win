using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using Spire.Doc;
using Word = Microsoft.Office.Interop.Word;


namespace macros_remover_win
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label4.Text = "";
            label5.Text = "";
            progressBar1.Visible = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //run
            label4.Text = "";
            label5.Text = "";
            listBox1.Items.Clear();
            button1.Enabled = false;
            button2.Enabled = false;
            checkBox2.Checked = true;
            Word.Application wordApp = new Word.Application();

            int percent = 1;
            int removed = 0;
            Console.WriteLine("------------------------------ Starting search for macroses... --------------------------");

            label4.Text = "Підготовка...";
            label4.Refresh();
            label5.Refresh();
            string arg = textBox1.Text;

            var filePaths = GetFiles(arg, "*.doc?");
            progressBar1.Visible = true;
            for (int i = 0; i < filePaths.Count; i++)
            {
                //Console.WriteLine(filePaths[i]);
                
                progressBar1.Value = percent;
                percent = 100 * i / filePaths.Count;
                string filePath = filePaths[i];

                label4.Text = "Шукаю в: " + filePath;
                label5.Text = percent + "%";
                label4.Refresh();
                label5.Refresh();
                listBox1.Refresh();
                progressBar1.Refresh();

                try
                {
                    var result = ProcessDocument(filePath, !checkBox1.Checked, wordApp);

                    if (result.MacrosDeleted)
                        removed++;

                    if (result.ContainsMacros)
                        listBox1.Items.Add(result.FilePath);

                    if (result.ExceptionOccured)
                        listBox1.Items.Add("Помилка в: " + result.FilePath);

                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                    wordApp.Quit();
                }

                Console.Write("\r{0}%   ", percent);
            }
            Console.WriteLine("------------------------------ All done -------------------------------");
            Console.WriteLine("Total docs fixed: " + removed);

            label4.Text = "Готово, всього очищено файлів: " + removed;
            label5.Text = "";
            progressBar1.Visible = false;
            button1.Enabled = true;
            button2.Enabled = true;
            wordApp.Quit();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private DocumentProcessResult ProcessDocument(string filePath, bool clearMacros, Word.Application wordApp)
        {
            
            var result = new DocumentProcessResult();
            result.FilePath = filePath;


            //Initialize a Document object
            using (Document document = new Document())
            {
                //Load the Word document
                try
                {
                    document.LoadFromFile(filePath);
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return result;
                }

                try
                {
                    //If the document contains macros, remove them from the document
                    if (document.IsContainMacro)
                    {
                        if (clearMacros)
                        {
                            try
                            {
                                document.ClearMacros();
                                document.SaveToFile(filePath + ".bak99", FileFormat.Docx);

                                File.Copy(filePath + ".bak99", filePath, true);
                                File.Delete(filePath + ".bak99");

                                result.MacrosDeleted = true;

                                Console.WriteLine("*** macros found and cleared: " + filePath);


                               
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                                result.ExceptionOccured = true;
                                result.MacrosDeleted = false;
                            }
                        }

                        result.ContainsMacros = true;
                    }
                    if (checkBox2.Checked) {
                        object missing = System.Reflection.Missing.Value;
                        //Word.Application wordApp = new Word.Application();
                        Word.Document aDoc = null;
                        object readOnly = false;
                        object isVisible = false;

                        wordApp.Visible = false;
                        object filename = filePath;
                        object saveAs = filePath + ".bak99";
                        object oTemplate = "";

                        aDoc = wordApp.Documents.Add(ref oTemplate, ref missing,
                                                     ref missing, ref missing);

                        aDoc = wordApp.Documents.Open(ref filename, ref missing,
                                                      ref readOnly, ref missing, ref missing,
                                                      ref missing, ref missing, ref missing,
                                                      ref missing, ref missing, ref missing,
                                                      ref isVisible, ref missing, ref missing,
                                                      ref missing, ref missing);

                        aDoc.Activate();
                        aDoc.set_AttachedTemplate(oTemplate);
                        aDoc.UpdateStyles();

                        aDoc.SaveAs(ref saveAs, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing);

                        aDoc.Close(ref missing, ref missing, ref missing);

                        File.Copy(filePath + ".bak99", filePath, true);
                        File.Delete(filePath + ".bak99");

                        result.MacrosDeleted = true;

                    }




                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return result;
                }
            }

            return result;
        }


        protected bool isDirectoryAccessible(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            if (!Directory.Exists(path))
            {
                return false;
            }

            try
            {
                Directory.EnumerateDirectories(path);
            }
            catch
            {
                return false;
            }

            return true;
        }

        protected bool isFileAccessible(string filename)
        {
            if (string.IsNullOrWhiteSpace(filename))
            {
                return false;
            }

            if (!File.Exists(filename))
            {
                return false;
            }

            try
            {
               File.GetAccessControl(filename, System.Security.AccessControl.AccessControlSections.Access);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public List<string> GetFiles(string path, string pattern)
        {
            List<string> fileList = new List<string>();
            List<string> directoryList = new List<string>();

            directoryList.Add(path);

            while (true)
            {
                if (directoryList.Count <= 0)
                {
                    break;
                }

                string directory = directoryList.First();
                directoryList.RemoveAt(0);

                if (!isDirectoryAccessible(directory))
                {
                    continue;
                }

                foreach (string item in Directory.EnumerateDirectories(directory, "*", SearchOption.TopDirectoryOnly))
                {
                    if (!isDirectoryAccessible(item))
                    {
                        continue;
                    }

                    directoryList.Add(item);
                }

                foreach (string item in Directory.EnumerateFiles(directory, pattern, SearchOption.TopDirectoryOnly))
                {
                    if (!isFileAccessible(item))
                    {
                        continue;
                    }

                    fileList.Add(item);
                }
            }

            return fileList;
        }

        private class DocumentProcessResult
        {
            public bool ContainsMacros;
            public bool MacrosDeleted;
            public bool ExceptionOccured;
            public string FilePath;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}

