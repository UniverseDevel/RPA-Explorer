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

            _rpaParser = rpaParser;
            
            comboBox1.Items.Add(RpaParser.Version.RPA_3_2);
            comboBox1.Items.Add(RpaParser.Version.RPA_3);
            comboBox1.Items.Add(RpaParser.Version.RPA_2);
            comboBox1.Items.Add(RpaParser.Version.RPA_1);

            if (_rpaParser._version == RpaParser.Version.Unknown)
            {
                comboBox1.SelectedItem = RpaParser.Version.RPA_3;
            }
            else
            {
                comboBox1.SelectedItem = _rpaParser._version;
            }

            textBox1.Text = _rpaParser._padding.ToString();
            textBox2.Text = _rpaParser._step.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _rpaParser._version = _rpaParser.CheckSupportedVersion((double) comboBox1.SelectedItem);
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch ((double) comboBox1.SelectedItem)
            {
                case RpaParser.Version.RPA_1:
                    textBox1.Enabled = false;
                    textBox2.Enabled = false;

                    textBox1.Text = "0";
                    textBox2.Text = "0";
                    break;
                case RpaParser.Version.RPA_2:
                    textBox2.Enabled = false;

                    textBox2.Text = "0";
                    break;
                default:
                    comboBox1.Enabled = true;
                    textBox1.Enabled = true;
                    textBox2.Enabled = true;
                    
                    comboBox1.SelectedItem = _rpaParser._version;
                    textBox1.Text = _rpaParser._padding.ToString();
                    textBox2.Text = _rpaParser._step.ToString();
                    break;
            }
        }
    }
}