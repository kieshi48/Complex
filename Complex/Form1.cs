using System;
using System.Windows.Forms;

namespace Complex
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<string> fileList = new List<string>();
        bool exceptions = false;

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["Delete"].Index)
            {
                dataGridView1.Rows.RemoveAt(e.RowIndex);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Word & Excel Documents|*.doc;*.docx;*.xls;*.xlsx";
            openFileDialog1.Title = "Select Word and Excel";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileList.Clear();
                textBox1.Text = Path.GetDirectoryName(openFileDialog1.FileName);
                foreach (string file in openFileDialog1.FileNames)
                {
                    fileList.Add(file);
                }
                MessageBox.Show($"Selected {fileList.Count} files");
            }
        }

        private async void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (fileList.Count == 0) throw new Exception("Please, select directory with Word files");

                var wordList = from files in fileList where (files.EndsWith(".docx") || files.EndsWith(".doc")) select files;

                if (wordList.Count() == 0) throw new Exception("Please, select directory with Word files");

                MessageBox.Show("Conversion procedure started");

                var converterToPDF = new ConverterToPDF();

                var task = converterToPDF.Converter(wordList);
                await task;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                exceptions = true;
            }
            finally
            {
                if (exceptions is false)
                {
                    MessageBox.Show("All Word documents in the folder have been converted to PDF.");
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var leftList = dataGridView1.Rows.Cast<DataGridViewRow>().Select(row => row.Cells[0].Value?.ToString()).ToList(); 
            var rightList = dataGridView1.Rows.Cast<DataGridViewRow>().Select(row => row.Cells[1].Value?.ToString()).ToList();
            

        }

        //private void button1_Click_2(object sender, EventArgs e)
        //{

        //}

        //private void button3_Click_2(object sender, EventArgs e)
        //{

        //}

        //private void button2_Click_2(object sender, EventArgs e)
        //{

        //}
    }
}