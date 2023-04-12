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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Word & Excel Documents|*.doc;*.docx;*.xls;*.xlsx";
            openFileDialog1.Title = "ֲûבונטעו פאיכû Word ט Excel";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = Path.GetDirectoryName(openFileDialog1.FileName);
                foreach (string file in openFileDialog1.FileNames)
                {
                    fileList.Add(file);
                }
                MessageBox.Show($"Selected {fileList.Count} files");
            }
        }
    }
}