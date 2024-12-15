using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BookcastLexicon
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Visible = true;
            textBox1.Multiline = true;
            textBox1.ScrollBars = ScrollBars.Vertical;
            label2.Text = "Bitte gib die Werte wie folgt an:\r\n\r\nBuchtitel-#Folgennummer-Zeitangabe-Schlagwort-Infos-Quellen\r\n\r\nBeispiel:\r\nStein der Weisen - #16 - 1:05:38 - Eule - Hedwig beißt Harry, habe Eulenverhalten analyisert - wikipedia.de/eulen";
            label2.Visible = true;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = true;
            button5.Text = "Übergeben";
            button5.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Visible = true;
            textBox1.Multiline = false;
            label2.Text = "Bitte lösche eine Zeile durch Angabe der folgenden Werte: \n\nBuchtitel-#Folgennummer-Zeitangabe-Schlagwort\n\nBeispiel:\nStein der Weisen - #16 - 1:05:38 - Eule";
            label2.Visible = true;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = true;
            button5.Text = "Löschen";
            button5.Visible = true;
        
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Visible = true;
            textBox1.Multiline = true;
            textBox1.ScrollBars = ScrollBars.Vertical;
            label2.Text = "Bitte gib ein Schlagwort ein, nach dem du suchen möchtest, z.B. Eule.";
            label2.Visible = true;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = true;
            button5.Text = "Suchen";
            button5.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            label2.Visible = false;
            button1.Visible = true;
            button2.Visible = true;
            button3.Visible = true;
            button4.Visible = false;
            button5.Visible = false;
            textBox1.Text = string.Empty;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            switch(button5.Text) 
            {
                case "Übergeben":
                    textBox1.Text = Program.CommitValues(splitStrings(textBox1.Text));
                    break;

                case "Löschen":
                    textBox1.Text = Program.DeleteValues(splitStrings(textBox1.Text));
                    break;

                case "Suchen":

                    textBox1.Text = Program.FindValue(textBox1.Text);
                    break;
            }
        }

        private string[] splitStrings(string input) 
        {
            return input.Split('-');
        }


    }
}
