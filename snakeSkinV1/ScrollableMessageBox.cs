using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace snakeSkinV1
{
    internal class ScrollableMessageBox
    {
        public static void Show(string message, string title)
        {
            // Create a new form
            Form form = new Form();
            form.Text = title;
            form.Width = 400;
            form.Height = 300;

            // Create a TextBox that will hold the message
            TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.ReadOnly = true;
            textBox.ScrollBars = ScrollBars.Vertical;
            textBox.Dock = DockStyle.Fill;
            textBox.Text = message;

            // Add the TextBox to the form
            form.Controls.Add(textBox);

            // Add an OK button to close the form
            Button buttonOk = new Button();
            buttonOk.Text = "OK";
            buttonOk.Dock = DockStyle.Bottom;
            buttonOk.Click += (sender, e) => form.Close();

            // Add the button to the form
            form.Controls.Add(buttonOk);

            // Display the form
            form.ShowDialog();
        }
    }
}
