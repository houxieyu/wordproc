using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
namespace wordproc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<FileInfo> FileList = new List<FileInfo>();
        Word.Application app = new Microsoft.Office.Interop.Word.Application();
        List<WordInfo> WordList = new List<WordInfo>();
        private void procfile(FileInfo fi)
        {
            try
            {
                Word.Document wordDoc = app.Documents.Open(fi.FullName);
                Microsoft.Office.Interop.Word.Range docRanger = wordDoc.Content;
                //页数
                int Pages = docRanger.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                //字数
                int Words = docRanger.ComputeStatistics(Word.WdStatistic.wdStatisticWords);
                //字符数（不计空格）
                int Characters = docRanger.ComputeStatistics(Word.WdStatistic.wdStatisticCharacters);
                //字符数（计空格）
                int Characterswithspaces = docRanger.ComputeStatistics(Word.WdStatistic.wdStatisticCharactersWithSpaces);
                //段落数
                int Paragraphs = docRanger.ComputeStatistics(Word.WdStatistic.wdStatisticParagraphs);
                //行数
                int Lines = docRanger.ComputeStatistics(Word.WdStatistic.wdStatisticLines);
                //MessageBox.Show(fi.Name+":"+ Words.ToString());
                WordList.Add(new WordInfo { fileName = fi.Name, filePath = fi.DirectoryName, wordCount=Words,
                    createTime = fi.CreationTime.ToString(), lastTime = fi.LastAccessTime.ToString() });

                wordDoc.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public void SaveToCSV() 
        {
            StreamWriter writer = new StreamWriter("word文档统计.csv",false, Encoding.UTF8);
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("文件名,文件路径,创建时间,最后修改时间,字数");
            foreach(WordInfo wi in WordList)
            {
                sb.AppendFormat("{0},{1},{2},{3},{4}\r\n",wi.fileName,wi.filePath,wi.createTime,wi.lastTime,wi.wordCount);
            }
            writer.Write(sb.ToString());
            writer.Close();
            MessageBox.Show("文档扫描完毕，共找到" + WordList.Count + "个word文档，请在\"word文档统计.csv\"中查看统计结果");
        }
        public void GetAllFiles(DirectoryInfo dir)
        {
            FileInfo[] allFile;
            try
            {
                allFile = dir.GetFiles();
            }
            catch (Exception) { return; }
            foreach (FileInfo fi in allFile)
            {
                label1.Text = fi.FullName;
                if (fi.Extension == ".doc" || fi.Extension == ".docx")
                {
                    procfile(fi);
                }
            }
            DirectoryInfo[] allDir = dir.GetDirectories();
            foreach (DirectoryInfo d in allDir)
            {
                GetAllFiles(d);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                GetAllFiles(new System.IO.DirectoryInfo(folderBrowserDialog1.SelectedPath));
                SaveToCSV();
            }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            app.Quit();
        }
    }
    public class WordInfo
    {
        public string fileName;
        public string filePath;
        public string createTime;
        public string lastTime;
        public int wordCount;
    }
}
