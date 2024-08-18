using System;
using System.Windows.Forms;

namespace snakeSkinV1
{
    public partial class BrowserFormT1 : Form
    {
        private DataGridView dataGridView;

        public BrowserFormT1()
        {
            // Initialize DataGridView
            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = true
            };

            // Add a column for text
            dataGridView.Columns.Add("TextColumn", "Text");

            // Add a column with buttons
            var buttonColumn = new DataGridViewButtonColumn
            {
                Name = "ButtonColumn",
                HeaderText = "Action",
                Text = "Click Me",
                UseColumnTextForButtonValue = true
            };
            dataGridView.Columns.Add(buttonColumn);

            // Fill DataGridView with some data
            for (int i = 0; i < 5; i++)
            {
                dataGridView.Rows.Add($"Row {i + 1}");
            }

            dataGridView.CellClick += DataGridView_CellClick;

            // Add DataGridView to the form
            this.Controls.Add(dataGridView);

            // Set form properties
            this.Width = 800;
            this.Height = 600;
            
            // Handle FormClosing event
            this.FormClosing += BrowserFormT1_FormClosing;
        }

        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView.Columns["ButtonColumn"].Index && e.RowIndex >= 0)
            {
                MessageBox.Show($"Button clicked in row {e.RowIndex + 1}!");
            }
        }

        private void BrowserFormT1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Prompt the user to confirm closing
            var result = MessageBox.Show("Are you sure you want to close?", "Confirm Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                // Cancel the close operation
                e.Cancel = true;
            }
        }

        // Method to add a new row
        public void AddRow(string text)
        {
            dataGridView.Rows.Add(text);
        }
    }
}
