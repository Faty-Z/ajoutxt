using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ajoutxt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Globals.ThisAddIn.X1();
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.X1();
        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            textBox3.Multiline = true;
            label1.Text = textBox3.Text;
            Globals.ThisAddIn.X1();
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.Multiline = true;
            label2.Text = textBox2.Text;
            Globals.ThisAddIn.X2();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Multiline = true;
            label3.Text = textBox1.Text;
            Globals.ThisAddIn.X3();
        }

        private void TextBox4_TextChanged(object sender, EventArgs e)
        {
            //textBox4.Multiline = true;
            //label4.Text= textBox4.Text;
            Globals.ThisAddIn.X4();
        }

        private void ListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.X5();
        }
    }
}
