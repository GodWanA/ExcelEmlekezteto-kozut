using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelEmlekezteto
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static bool _Talcara { get; set; } = false;

        protected override void OnStartup(StartupEventArgs e)
        {
            TalcaIkon.StartUp();
            base.OnStartup(e);

            foreach (string item in e.Args)
            {
                switch (item.ToLower())
                {
                    case "-minimized":
                        App._Talcara = true;
                        break;
                }
            }
        }
    }
}
