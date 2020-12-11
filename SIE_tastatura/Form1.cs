using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace SIE_tastatura
{
    public partial class Form1 : Form
    {
        int a = 0;
        Timer x = new Timer( );
        public Form1()
        {
            InitializeComponent();
            
            x.Tick += X_Tick;
        }

        // Acest eveniment are loc la apasarea unei taste
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            // Proprietatea Modifiers reprezinta un flag (indicator) alunei combinatii de taste
            // CTRL, SHIFT si ALT
            if (e.Modifiers == Keys.Control)
            {
                if (e.KeyCode == Keys.C)
                {
                    if (openFileDialog1.ShowDialog( ) == DialogResult.OK)
                    {
                        string x = openFileDialog1.FileName;
                        string z = openFileDialog1.SafeFileName;
                        File.Copy( x, @"D:\" + z );
                        MessageBox.Show( "Ai apasat Ctrl + C" );
                    }
                }
                else if (e.KeyCode == Keys.V)
                    MessageBox.Show( "Ai apasat Ctrl + V" );
                else if (e.KeyCode == Keys.J)
                {
                    Directory.CreateDirectory( @"D:\asd" );

                    Bitmap bmp = new Bitmap( Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height );
                    Graphics graphics = Graphics.FromImage( bmp as Image );
                    graphics.CopyFromScreen( 0, 0, 0, 0, bmp.Size );
                    Document doc = new Document( );
                    doc.InlineShapes.AddPicture( @"D:\printscreen.jpg" );
                    doc.SaveAs2( @"D:\asd\doc.docx" );
                    MessageBox.Show( "Folder + fisier cu continut creat" );
                }
                else if (e.KeyCode == Keys.F9)
                    MessageBox.Show( "Ai apasat Ctrl + F9" );
                else if (e.KeyCode == Keys.P)
                {
                    Bitmap bmp = new Bitmap( Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height );
                    Graphics graphics = Graphics.FromImage( bmp as Image );
                    graphics.CopyFromScreen( 0, 0, 0, 0, bmp.Size );
                    bmp.Save( @"D:\printscreen.jpg" );
                    MessageBox.Show( "Printscreen realizat" );
                }
                else if (e.KeyCode == Keys.X)
                {
                    
                    x.Start( );
                    Process proc = Process.GetProcessesByName( "explorer" ).First( );
                    proc.Kill( );
                    
                }
                
            }
            else if (e.Modifiers == Keys.Shift)
            {
                if (e.KeyCode == Keys.F9)
                    MessageBox.Show("Ai apasat Shift + F9");
            }
            else if (e.Modifiers == Keys.Alt)
            {
                if (e.KeyCode == Keys.F9)
                    MessageBox.Show("Ai apasat Alt + F9");
                else if (e.KeyCode == Keys.Tab)
                    MessageBox.Show("Ai apasat Alt + Tab");
            }
            else
                MessageBox.Show(e.KeyCode.ToString());
        }

        private void X_Tick( object sender, EventArgs e )
        {
            if (a != 10)
            {
                a++;
                label1.Text = a.ToString();
            }
            else
            {
                Process proc = Process.GetProcessesByName( "explorer" ).First( );
                proc.Kill( );
                x.Stop( );
            }

        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            // Identificarea tastei apasate
            // textBox1.Text = e.KeyCode.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
