using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace ExcelEmlekezteto
{
    public static class TalcaIkon
    {

        static public NotifyIcon _TalcaIkon { get; set; } = new NotifyIcon();

        /// <summary>
        /// Egy értesítési ablakban megjeleníti az összeállított állományt.
        /// </summary>
        /// <param name="cim">A címsor szövege</param>
        /// <param name="uzenet">Az értesítés szövege</param>
        /// /// <param name="ablak">Az ablak amit visszanyit</param>
        static public void Ertesites(string cim, string uzenet, Window ablak = null)
        {
            _TalcaIkon.BalloonTipTitle = cim;
            _TalcaIkon.BalloonTipText = uzenet;

            _TalcaIkon.ShowBalloonTip(30);

            if (ablak != null)
            {
                ablak.Show();
            }
        }

        public static void StartUp()
        {
            TalcaIkon._TalcaIkon.Icon = Properties.Resources.icon1;
            TalcaIkon._TalcaIkon.Visible = true;
        }
    }
}
