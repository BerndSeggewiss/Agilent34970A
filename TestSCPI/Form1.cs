// Testing device 34970A for 4-wire resistor measurement.
// Version V3:
// 18.10.2024:
// - Read all 30 resistors values and display them in the textboxes.
// - Display the resistor names in the labels.
// - Display the HW-ID in the label.
// - Make resistor values visible only if the checkbox is checked.
// Version V3:
// 22.10.2024:
// - Add DataGridView to display the resistor values in a table.
// - Add a numericUpDown control to select the table row.
// - Highlight the selected row in the DataGridView.

using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace TestSCPI
{
    public partial class Form1 : Form
    {
        private DataGridView dataGridView;
        private TextBox[] textBoxes;
        private CheckBox[] checkBoxes;
        private string[] ResistorlabelNames = {"R101", "R102", "R103", "R104", "R105", "R106", "R107", 
            "R108", "R109", "R110", "R201", "R202", "R203", "R204", "R205", "R206", "R207", "R208", "R209",
            "R210", "R301", "R302", "R303", "R304", "R305", "R306", "R307", "R308", "R309", "R310" };
        private string[] TemperatureStepPoints = {"T = 20", "T = 40", "T = 60", "T = 80", "T = 90"};

        public Form1()
        {
            InitializeComponent();
            InitializeTextBoxes();
            InitializeCheckBoxes();
            InitializeDataGridView();
        }

        private void InitializeNumericUPDown()
        {
            numericUpDownTableRow.Minimum = 1;
            numericUpDownTableRow.Maximum = 6;
        }

        private void InitializeDataGridView()
        {
            Point panelLocation = dataPanel.Location;
            int dataGridViewX = panelLocation.X;
            int dataGridViewY = this.ClientSize.Height - 200;
            dataGridView = new DataGridView
            {
                Location = new Point(dataGridViewX, dataGridViewY),
                Size = new Size(dataPanel.Width, 170),
                ColumnCount = ResistorlabelNames.Length,
                RowCount = TemperatureStepPoints.Length,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                GridColor = Color.Black
            };
            for (int i = 0; i < ResistorlabelNames.Length; i++)
            {
                dataGridView.Columns[i].Name = ResistorlabelNames[i];
            }
            dataGridView.RowHeadersWidth = 71;
            for (int i = 0; i < TemperatureStepPoints.Length; i++)
            {
                dataGridView.Rows[i].HeaderCell.Value = TemperatureStepPoints[i];
            }
            this.Controls.Add(dataGridView);
            dataGridView.Width = dataPanel.Width + controlPanel.Width + 10;
            controlPanel.Location = new Point(dataPanel.Location.X + dataPanel.Width + 10, dataPanel.Location.Y);
            controlPanel.Height =  dataPanel.Height;
        }

        private void InitializeTextBoxes()
        {
            textBoxes = new TextBox[] { textBoxResistor1, textBoxResistor2, textBoxResistor3, textBoxResistor4,
                textBoxResistor5, textBoxResistor6, textBoxResistor7, textBoxResistor8, textBoxResistor9, 
                textBoxResistor10, textBoxResistor11, textBoxResistor12, textBoxResistor13, textBoxResistor14,
                textBoxResistor15, textBoxResistor16, textBoxResistor17, textBoxResistor18, textBoxResistor19, 
                textBoxResistor20, textBoxResistor21, textBoxResistor22, textBoxResistor23, textBoxResistor24,
                textBoxResistor25, textBoxResistor26, textBoxResistor27, textBoxResistor28, textBoxResistor29,
                textBoxResistor30 };
        }

        private void InitializeCheckBoxes()
        {
            checkBoxes = new CheckBox[] { checkBoxResistor1, checkBoxResistor2, checkBoxResistor3, checkBoxResistor4,
                checkBoxResistor5, checkBoxResistor6, checkBoxResistor7, checkBoxResistor8, checkBoxResistor9, 
                checkBoxResistor10, checkBoxResistor11, checkBoxResistor12, checkBoxResistor13, checkBoxResistor14, 
                checkBoxResistor15, checkBoxResistor16, checkBoxResistor17, checkBoxResistor18, checkBoxResistor19, 
                checkBoxResistor20, checkBoxResistor21, checkBoxResistor22, checkBoxResistor23, checkBoxResistor24, 
                checkBoxResistor25, checkBoxResistor26, checkBoxResistor27, checkBoxResistor28, checkBoxResistor29,
                checkBoxResistor30};
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            labelDataPanelInfo.Text = "Values 4-Wire connected";
            labelDataPanelInfo.Font = new Font("Arial", 10, FontStyle.Bold);
            labelInfoTableRow.Text = "Table Row";
            labelAppDesignation.Text = "Agilent 34970A 4-wire resistor measurement Version 3";
            labelAppDesignation.Font = new Font("Arial", 12, FontStyle.Bold);
            labelResultTableInfo.Text = "Resistor Result Table";
            labelResultTableInfo.Font = new Font("Arial", 10, FontStyle.Bold);
            getDevHWID();
        }

        private void getDevHWID()
        {
            try
            {
                using (var session = (Ivi.Visa.IMessageBasedSession)Ivi.Visa.GlobalResourceManager.Open("ASRL5::INSTR"))
                {
                    session.FormattedIO.WriteLine("*IDN?");
                    string hwID = session.FormattedIO.ReadString();
                    labelDeviceInfo.Text = "Device HW-ID: ";
                    labelDeviceInfo.Font = new Font("Arial", 10);
                    labelHW_ID.Text = hwID;
                    labelHW_ID.Font = new Font("Arial", 10);
                    labelHW_ID.BackColor = Color.LightGreen;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading HW-ID, check connection... " + ex.Message);
            }
        }

        private void readResistorValues()
        {
            try
            {
                using (var session = (Ivi.Visa.IMessageBasedSession)Ivi.Visa.GlobalResourceManager.Open("ASRL5::INSTR"))
                {
                    session.FormattedIO.WriteLine(":CONFigure:FRESistance AUTO,MAX,(@101:110,201:210,301:310)");
                    session.FormattedIO.WriteLine(":TRIGger:SOURce TIMer");
                    session.FormattedIO.WriteLine(":TRIGger:TIMer 20");
                    session.FormattedIO.WriteLine(":TRIGger:COUNt 1");
                    session.FormattedIO.WriteLine(":INITiate");
                    session.FormattedIO.WriteLine(":FETCh?");
                    string readings = session.FormattedIO.ReadString();
                    string[] values = readings.Split(',');

                    UpdateTextBoxesAndLabels(values);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading resistor values... " + ex.Message);
            }
        }

        private void UpdateTextBoxesAndLabels(string[] values)
        {
            for (int i = 0; i < values.Length && i < textBoxes.Length; i++)
            {
                string trimmedValue = values[i].Trim();
                double parsedValue = double.Parse(trimmedValue, NumberStyles.Float, CultureInfo.InvariantCulture);
                string formattedValue = string.Format(CultureInfo.GetCultureInfo("de-DE"), 
                    "{0:0,0.0}", parsedValue).Replace(".", string.Empty);
                if (i < textBoxes.Length && !checkBoxes[i].Checked)
                {
                    textBoxes[i].Text = string.Empty;
                }
                else
                {
                    textBoxes[i].Text = formattedValue;
                    textBoxes[i].BackColor = Color.LightGoldenrodYellow;
                }
                textBoxResistor1.SelectionStart = 0;
                if (i < ResistorlabelNames.Length)
                {
                    string labelName = "Rlabel" + (i + 1);
                    Control[] foundControls = this.Controls.Find(labelName, true);
                    if (foundControls.Length > 0 && foundControls[0] is Label label)
                    {
                        label.Text = ResistorlabelNames[i];
                    }
                }
            }
            UpdateDataGridViewFromTextBoxes();
        }

        private void HighLightRow(int rowIndex)
        {
            if (dataGridView.Rows.Count > 0)
            {
                dataGridView.Rows[rowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                Timer timer = new Timer();
                timer.Interval = 3000;
                timer.Tick += (sender, e) =>
                {
                    dataGridView.Rows[rowIndex].DefaultCellStyle.BackColor = Color.White;
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
            }
        }

        private void UpdateDataGridViewFromTextBoxes()
        {
            int selectedRow = (int)numericUpDownTableRow.Value - 1;
            if (dataGridView.Rows.Count > 0)
            {
                for (int i = 0; i < textBoxes.Length && i < dataGridView.Columns.Count; i++)
                {
                    dataGridView.Rows[selectedRow].Cells[i].Value = textBoxes[i].Text;
                }
            }
        }

        private void buttonReadValues_Click(object sender, EventArgs e)
        {
            int selectedRow = (int)numericUpDownTableRow.Value - 1;
            HighLightRow(selectedRow);
            readResistorValues();
        }

        private void NumericDropDown_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownTableRow.Value < 1)
            {
                numericUpDownTableRow.Value = 1;
            }
            else if (numericUpDownTableRow.Value > 5)
            {
                numericUpDownTableRow.Value = 5;
            }
        }
    }
}
