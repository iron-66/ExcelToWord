using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToWord
{
    public partial class CustomMessageBox : Form
    {
        public CustomMessageBox()
        {
            InitializeComponent();

            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            // Таймер автоматического закрытия окна
            Timer timer = new Timer();
            timer.Interval = 4000; // 4 секунды
            timer.Tick += (sender, e) =>
            {
                Close();
                timer.Stop();
                timer.Dispose();
            };
            timer.Start();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
