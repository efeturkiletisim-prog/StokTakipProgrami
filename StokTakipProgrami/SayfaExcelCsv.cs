using System;
using System.Drawing;
using System.Windows.Forms;

namespace StokTakipProgrami
{
    public class SayfaExcelCsv : UserControl
    {
        public SayfaExcelCsv()
        {
            this.BackColor = Color.FromArgb(33, 33, 36);
            Label lbl = new Label
            {
                Text = "📊 TOPLU EXCEL / CSV AKTARIM ",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lbl);
        }
    }
}