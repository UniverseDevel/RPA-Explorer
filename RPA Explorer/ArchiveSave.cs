using System;
using System.Windows.Forms;
using RPA_Parser;

namespace RPA_Explorer
{
    public partial class ArchiveSave : Form
    {
        private RpaParser _rpaParser;
        public ArchiveSave(RpaParser rpaParser)
        {
            InitializeComponent();

            comboBox1.SelectedItem = rpaParser._version.ToString();
            textBox1.Text = rpaParser._padding.ToString();
            textBox2.Text = rpaParser._step.ToString();

            _rpaParser = rpaParser;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _rpaParser._version = RpaParser.CheckSupportedVersion(Convert.ToDouble(comboBox1.SelectedItem));
                _rpaParser._padding = Convert.ToInt32(textBox1.Text);
                _rpaParser._step = Convert.ToInt64(textBox2.Text);
                _rpaParser.optionsConfirmed = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Invalid values", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}