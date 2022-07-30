using System;
using System.Windows.Forms;
using RPA_Parser;

namespace RPA_Explorer
{
    public partial class ArchiveSave : Form
    {
        private readonly RpaParser _rpaParser;
        public ArchiveSave(RpaParser rpaParser)
        {
            InitializeComponent();

            _rpaParser = rpaParser;

            LoadTexts();
            
            comboBox1.Items.Add(RpaParser.Version.RPA_3_2);
            comboBox1.Items.Add(RpaParser.Version.RPA_3);
            comboBox1.Items.Add(RpaParser.Version.RPA_2);
            comboBox1.Items.Add(RpaParser.Version.RPA_1);

            comboBox1.SelectedItem = _rpaParser.CheckVersion(_rpaParser.ArchiveVersion, RpaParser.Version.Unknown) ? RpaParser.Version.RPA_3 : _rpaParser.ArchiveVersion;

            textBox1.Text = _rpaParser.Padding.ToString();
            textBox2.Text = _rpaParser.ObfuscationKey.ToString();

        }

        private void LoadTexts()
        {
            Text = RpaExplorer.GetText("Archive_save_title");
            label1.Text = RpaExplorer.GetText("Archive_save_version");
            label2.Text = RpaExplorer.GetText("Archive_save_padding");
            label3.Text = RpaExplorer.GetText("Archive_save_obfuscationkey");
            button1.Text = RpaExplorer.GetText("Archive_save_continue");
            button2.Text = RpaExplorer.GetText("Archive_save_cancel");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _rpaParser.ArchiveVersion = _rpaParser.CheckSupportedVersion((double) comboBox1.SelectedItem);
                _rpaParser.Padding = Convert.ToInt32(textBox1.Text);
                _rpaParser.ObfuscationKey = Convert.ToInt64(textBox2.Text);
                _rpaParser.OptionsConfirmed = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    RpaExplorer.GetText("Invalid_values"), MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    textBox1.Text = 0.ToString();
                    textBox2.Text = 0.ToString();
                    break;
                case RpaParser.Version.RPA_2:
                    textBox1.Enabled = true;
                    textBox2.Enabled = false;

                    textBox1.Text = _rpaParser.Padding.ToString();
                    textBox2.Text = 0.ToString();
                    break;
                default:
                    comboBox1.Enabled = true;
                    textBox1.Enabled = true;
                    textBox2.Enabled = true;
                    
                    textBox1.Text = _rpaParser.Padding.ToString();
                    textBox2.Text = _rpaParser.ObfuscationKey.ToString();
                    break;
            }
        }
    }
}