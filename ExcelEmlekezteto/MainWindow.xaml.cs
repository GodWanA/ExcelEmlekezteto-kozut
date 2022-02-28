using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.CodeDom;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace ExcelEmlekezteto
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string _File
        {
            get
            {
                return this._file;
            }
            set
            {
                this._file = value;
                this.fileSystem = null;

                FileInfo file = new FileInfo(this._File);

                if (file.Exists && file.Extension == ".xlsx")
                {
                    this.fileSystem = new FileSystemWatcher(file.DirectoryName, file.Name);
                    this.fileSystem.EnableRaisingEvents = true;

                    this.fileSystem.Changed += this.FileSystem_Changed;
                    this.fileSystem.Deleted += this.FileSystem_Deleted;
                    this.torolve = false;
                }

                file = null;
            }
        }

        private static DispatcherTimer emlekezteto = new DispatcherTimer();

        private bool torolve = false;

        private void FileSystem_Deleted(object sender, FileSystemEventArgs e)
        {
            if (this.IsLoaded && !this.torolve)
            {
                this.torolve = true;
                Action a = () =>
                {
                    MessageBox.Show("A \"" + this.textBox_file.Text + "\" állományt törölték.\r\nKérem jelölje ki újra!", "Figyelem!", MessageBoxButton.OK, MessageBoxImage.Warning);
                };
                Dispatcher.BeginInvoke(a);
            }
        }

        public string _Starthely { get; set; } = System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("ExcelEmlekezteto.exe", "");
        public DataTable dt { get; set; } = new DataTable();
        public int _Hibas { get; set; } = 0;
        public int _Lejart { get; set; } = 0;
        public int _LeFogJarni { get; set; } = 0;
        public int _PreferaltOszlop { get; set; }

        private FileSystemWatcher fileSystem;
        private string _file;
        private string bealitas;
        private bool betoltesAlatt;

        public MainWindow()
        {
            InitializeComponent();

            TalcaIkon._TalcaIkon.Text = this.Title;
            TalcaIkon._TalcaIkon.Click += this._TalcaIkon_Click;
            TalcaIkon._TalcaIkon.BalloonTipClicked += this._TalcaIkon_BalloonTipClicked;

            MainWindow.emlekezteto.Interval = TimeSpan.FromMinutes(10);
            MainWindow.emlekezteto.Tick += this.Emlekezteto_Tick;
            MainWindow.emlekezteto.Start();
        }

        private void Emlekezteto_Tick(object sender, EventArgs e)
        {
            this.GridAdatotSzinez();
            if (
                this._Lejart > 0 ||
                this._LeFogJarni > 0
                )
            {
                new Thread(delegate ()
                {
                    Action a = () => this.Frissul();

                    Dispatcher.BeginInvoke(a);
                }).Start();
            }
            MainWindow.emlekezteto.Interval = TimeSpan.FromDays(1);
        }

        private void _TalcaIkon_BalloonTipClicked(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Normal;
            this.Topmost = true;
            this.Topmost = false;
        }

        private void _TalcaIkon_Click(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Normal;
            this.Topmost = true;
            this.Topmost = false;
        }

        private void stackpanel_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Note that you can have more than one file.
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                foreach (string file in files)
                {
                    if (IsXlsx(file)) break;
                }

                this.SafeOpenFile();
                this.BealitastMent();
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "Excel munkafüzet (*.xlsx)|*.xlsx|Minden fájl (*.*)|*.*";
            openFileDialog.FilterIndex = 0;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "Munkafüzet kiválasztása";

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (this.IsXlsx(openFileDialog.FileName))
                {
                    SafeOpenFile();
                    BealitastMent();
                }
                else MessageBox.Show("Hibás állomány! \r\n" + openFileDialog.FileName, "Hiba a fájl kiterjesztésével!");
            }
        }

        /// <summary>
        /// Megnézi, hogy a kiválasztott fájl xlsx-e, ha igen a fájlt be is állítja a kijelzett érték helyén.
        /// </summary>
        /// <param name="file">A fájl elérési útja</param>
        /// <returns>Boolean</returns>
        private bool IsXlsx(string file)
        {
            if (!File.Exists(file)) return false;

            FileInfo fileInfo = new FileInfo(file);

            if (fileInfo.Extension == ".xlsx")
            {
                this.textBox_file.Text = file;
                this._File = file;
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Ellenőrzi, hogy a fájl megnyitható-e, ha igen felolvassa az excel állományt.
        /// </summary>
        private async void SafeOpenFile()
        {
            this.dt = await OpenFile();

            if (this.comboBox_oszlop.SelectedIndex != -1) this._PreferaltOszlop = this.comboBox_oszlop.SelectedIndex;
            this.comboBox_oszlop.Items.Clear();

            for (int i = 0; i < this.dt.Columns.Count; i++)
            {
                this.comboBox_oszlop.Items.Add(new ComboBoxItem().Content = i + 1 + " - " + dt.Rows[0][i]);
            }

            this.dataGrid_demo.ItemsSource = dt.DefaultView;

            if (this.comboBox_oszlop.Items.Count > 0)
            {
                if (this._PreferaltOszlop > -1 && this.comboBox_oszlop.Items.Count > this._PreferaltOszlop) this.comboBox_oszlop.SelectedIndex = this._PreferaltOszlop;
                else this.comboBox_oszlop.SelectedIndex = 0;
            }
            this.GridAdatotSzinez();
        }


        /// <summary>
        /// Aszinkron módon olvassa fel az excel állományt;
        /// </summary>
        private Task<DataTable> OpenFile()
        {
            DataTable tmp = new DataTable();

            if (this._File != null && this._File != "" && this.IsXlsx(this._File))
            {
                while (true)
                {
                    try
                    {
                        using (var stream = File.Open(this._File, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            tmp = reader.AsDataSet().Tables[0];
                            tmp.Columns.Add(new DataColumn());
                            tmp.Rows[0][tmp.Columns.Count - 1] = "ÜZENET";
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(1000);
                    }
                }
            }

            foreach (DataRow row in tmp.Rows)
            {
                for (int i = 0; i < tmp.Columns.Count; i++)
                {
                    if (row[i] is DateTime)
                    {
                        row[i] = ((DateTime)row[i]).ToString("yyyy.MM.dd");
                    }
                }
            }

            return Task.FromResult<DataTable>(tmp);
        }

        /// <summary>
        /// Befesti a grid adatait;
        /// </summary>
        private void GridAdatotSzinez()
        {
            bool elso = true;
            this._Hibas = 0;
            this._LeFogJarni = 0;
            this._Lejart = 0;

            foreach (DataRowView dr in dataGrid_demo.Items)
            {
                this.dataGrid_demo.UpdateLayout();
                DataGridRow dgr = this.dataGrid_demo.ItemContainerGenerator.ContainerFromItem(dr) as DataGridRow;
                dgr.ToolTip = null;
                //dgr.Background = Brushes.White;
                //dgr.Foreground = Brushes.Black;
                //dgr.FontWeight = FontWeight.FromOpenTypeWeight(600);

                if (elso)
                {
                    dgr.Background = Brushes.Green;
                    dgr.FontWeight = FontWeight.FromOpenTypeWeight(900);
                    dgr.Foreground = Brushes.White;
                    elso = false;
                }
                else
                {
                    if (this.comboBox_oszlop.SelectedIndex > -1)
                    {
                        string v = dr[this.comboBox_oszlop.SelectedIndex].ToString();

                        if (DateTime.TryParseExact(v, "yyyy.MM.dd", null, System.Globalization.DateTimeStyles.None, out DateTime tmp))
                        {
                            tmp = this.Lejarat(tmp);

                            int lejart = (int)(DateTime.Today - tmp.Date).TotalDays;

                            if (lejart == 0) dr[dt.Columns.Count - 1] = "Ma ját le a műszaki";
                            else if (lejart < 0) dr[dt.Columns.Count - 1] = (lejart * -1) + " nap mulva jár le a múszaki";
                            else dr[dt.Columns.Count - 1] = lejart + " napja járt le a műszaki";

                            dgr.Background = Brushes.White;
                            dgr.Foreground = Brushes.Black;
                            dgr.FontWeight = FontWeight.FromOpenTypeWeight(775);

                            switch (this.IsLejart(tmp.Date))
                            {
                                case 1:
                                    dgr.Background = Brushes.Orange;
                                    dgr.Foreground = Brushes.Black;
                                    this._LeFogJarni++;
                                    break;
                                case 2:
                                    dgr.Background = Brushes.Red;
                                    dgr.Foreground = Brushes.White;
                                    this._Lejart++;
                                    break;
                                default:
                                    dgr.FontWeight = FontWeight.FromOpenTypeWeight(400);
                                    break;
                            }
                        }
                        else if (dr[comboBox_oszlop.SelectedIndex].GetType() != Type.GetType("System.DateTime"))
                        {
                            dgr.Background = Brushes.Yellow;
                            dgr.Foreground = Brushes.Black;
                            dgr.FontWeight = FontWeight.FromOpenTypeWeight(650);
                            this._Hibas++;
                            dr[dt.Columns.Count - 1] = "Hibás, vagy hiányos a vizsgált adat";
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Eldönti hogy a kapott dátum lejárt-e már, vagy sem.
        /// </summary>
        /// <param name="date">kérdéses dátum</param>
        /// <returns>
        /// 0: Esetén a dátum nem járt még le<br/>
        /// 1: Esetén a már megadott időn belül le fog járni<br/>
        /// 2: Esetén már lejárt<br/>
        /// </returns>
        private int IsLejart(DateTime date)
        {
            DateTime elore = this.Figyelmeztet(DateTime.Today);

            if (elore.Date >= date.Date)
            {
                if (DateTime.Today < date.Date) return 1;
                else return 2;
            }
            else return 0;
        }

        /// <summary>
        /// Kiszámolja a lejárta időpontját.
        /// </summary>
        /// <param name="date">A dátum amihez hozzáadom a lejárat időtartamát</param>
        /// <returns>Datetime</returns>
        private DateTime Lejarat(DateTime date)
        {
            int tmp;

            if (int.TryParse(this.textBox_lejart_ev.Text, out tmp)) date = date.AddYears(tmp);
            if (int.TryParse(this.textBox_lejart_honap.Text, out tmp)) date = date.AddMonths(tmp);
            if (int.TryParse(this.textBox_lejart_nap.Text, out tmp)) date = date.AddDays(tmp);

            return date;
        }

        /// <summary>
        /// Kiszámolja a figyelmeztetés időpontját.
        /// </summary>
        /// <param name="date">A dátum amihez hozzáadom a lejárat időtartamát</param>
        /// <returns>Datetime</returns>
        private DateTime Figyelmeztet(DateTime date)
        {
            int tmp;

            if (int.TryParse(this.textBox_figyelmeztet_ev.Text, out tmp)) date = date.AddYears(tmp);
            if (int.TryParse(this.textBox_figyelmeztet_honap.Text, out tmp)) date = date.AddMonths(tmp);
            if (int.TryParse(this.textBox_figyelmeztet_nap.Text, out tmp)) date = date.AddDays(tmp);

            return date;
        }


        private void textBox_file_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.textBox_file.IsFocused)
            {
                this._File = this.textBox_file.Text;
                this.SafeOpenFile();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.BeallitastBetolt();
            if (App._Talcara) this.WindowState = WindowState.Minimized;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.BealitastMent();

            //e.Cancel = true;
            TalcaIkon._TalcaIkon.Visible = false;
        }

        private void comboBox_oszlop_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.IsLoaded && !this.betoltesAlatt)
            {
                //this.RegistryIras("Oszlop", this.comboBox_oszlop.SelectedIndex);
                this.BealitastMent();
                this.GridAdatotSzinez();

                if (this.comboBox_oszlop.SelectedIndex > -1) this._PreferaltOszlop = this.comboBox_oszlop.SelectedIndex;
            }
        }

        private void textBox_lejart_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.IsLoaded)
            {
                //this.RegistryIras("Lejart_ertek", this.textBox_lejart.Text);
                this.BealitastMent();
                this.GridAdatotSzinez();
            }
        }

        private void textBox_figyelmeztet_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.IsLoaded)
            {
                //this.RegistryIras("Figyelmeztet_ertek", this.textBox_figyelmeztet.Text);
                this.BealitastMent();
                this.GridAdatotSzinez();
            }
        }

        private void textBox_email_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.IsLoaded && this.textBox_email.IsFocused)
            {
                this.BealitastMent();
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            FileInfo file = new FileInfo(this.textBox_file.Text);

            ProcessStartInfo processStartInfo = new ProcessStartInfo();
            processStartInfo.WorkingDirectory = file.DirectoryName;
            processStartInfo.FileName = file.FullName;

            Process.Start(processStartInfo);

            file = null;
            processStartInfo = null;
        }

        private bool olvas = false;

        private void FileSystem_Changed(object sender, FileSystemEventArgs e)
        {
            if (!this.olvas)
            {
                this.olvas = true;

                Action action = () =>
                {
                    if (this.IsLoaded)
                    {
                        int index = this.comboBox_oszlop.SelectedIndex;
                        this.SafeOpenFile();
                        this.comboBox_oszlop.SelectedIndex = index;

                        if (
                            (this.IsLoaded) &&
                            (this._Hibas > 0 || this._LeFogJarni > 0 || this._Lejart > 0)
                        )
                        {
                            this.Frissul();
                            MainWindow.emlekezteto.Stop();
                            MainWindow.emlekezteto.Start();
                            this.olvas = false;
                        }
                    }
                };

                Dispatcher.BeginInvoke(action);
            }
        }

        /// <summary>
        /// Menti az ablakra vonatkozó adatokat egy ini fájlba.
        /// </summary>
        private void BealitastMent()
        {
            if (this.IsLoaded && !this.betoltesAlatt && this.comboBox_oszlop.SelectedIndex > -1)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("top = \"" + this.Top + "\"");
                sb.AppendLine("left = \"" + this.Left + "\"");
                sb.AppendLine("width = \"" + this.ActualWidth + "\"");
                sb.AppendLine("height = \"" + this.ActualHeight + "\"");
                sb.AppendLine("windiwstate = \"" + this.WindowState + "\"");
                sb.AppendLine("file = \"" + this._file + "\"");
                sb.AppendLine("oszlop = \"" + this.comboBox_oszlop.SelectedIndex + "\"");
                sb.AppendLine("email = \"" + this.textBox_email.Text.Replace("\r\n", ";") + "\"");
                sb.AppendLine("lejart_ev = \"" + this.textBox_lejart_ev.Text + "\"");
                sb.AppendLine("lejart_honap = \"" + this.textBox_lejart_honap.Text + "\"");
                sb.AppendLine("lejart_nap = \"" + this.textBox_lejart_nap.Text + "\"");
                sb.AppendLine("figyelmeztet_ev = \"" + this.textBox_figyelmeztet_ev.Text + "\"");
                sb.AppendLine("figyelmeztet_honap = \"" + this.textBox_figyelmeztet_honap.Text + "\"");
                sb.AppendLine("figyelmeztet_nap = \"" + this.textBox_figyelmeztet_nap.Text + "\"");
                sb.AppendLine("inditas_windowssal = \"" + this.checkBox.IsChecked + "\"");

                using (StreamWriter sw = new StreamWriter(this._Starthely + "config.ini"))
                {
                    sw.Write(sb.ToString());

                    sw.Flush();
                    sw.Close();
                }
            }
        }

        /// <summary>
        /// Értesítést küld a windowsnak, illetve levelet a megadott címekre.
        /// </summary>
        private void Frissul()
        {
            TalcaIkon.Ertesites("Figyelem",
                                    "Lejart adat: " + this._Lejart + " db.\r\n" +
                                    "Lejáró adat: " + this._LeFogJarni + " db.\r\n" +
                                    "Hibás adat: " + this._Hibas + " db."
                                    , this
                                );
            EmailKuldes();
        }

        /// <summary>
        /// Visszatölti az adatokat a mentérhez is használt ini fájlból
        /// </summary>
        private void BeallitastBetolt()
        {
            this.betoltesAlatt = true;
            if (File.Exists(this._Starthely + "config.ini"))
            {
                this.bealitas = File.ReadAllText(this._Starthely + "config.ini");
                string[] adatok = File.ReadAllLines(this._Starthely + "config.ini");

                foreach (string adat in adatok)
                {
                    string[] s = adat.Split(new string[] { " = ", " =", "= ", "=" }, StringSplitOptions.None);
                    if (s.Length == 2)
                    {
                        char[] szemet = new char[] { '\"', '\n', '\r', '\'', ' ', '\t' };

                        string kulcs = s[0].TrimStart(szemet).TrimEnd(szemet);
                        string ertek = s[1].TrimStart(szemet).TrimEnd(szemet);

                        switch (kulcs.ToLower())
                        {
                            case "top":
                                if (Double.TryParse(ertek, out double tmp_d))
                                {
                                    if (tmp_d < 0) this.Top = 0;
                                    else if (tmp_d > SystemParameters.WorkArea.Height) this.Top = SystemParameters.WorkArea.Height - this.ActualHeight;
                                    else this.Top = tmp_d;
                                }
                                break;
                            case "left":
                                if (Double.TryParse(ertek, out tmp_d))
                                {
                                    if (tmp_d < 0) this.Left = 0;
                                    else if (tmp_d > SystemParameters.VirtualScreenWidth) this.Left = SystemParameters.VirtualScreenWidth - this.ActualWidth;
                                    else this.Left = tmp_d;
                                }
                                break;
                            case "width":
                                if (Double.TryParse(ertek, out tmp_d))
                                {
                                    if (tmp_d < this.MinWidth) this.Width = this.MinWidth;
                                    else if (tmp_d > SystemParameters.VirtualScreenWidth) this.Width = SystemParameters.VirtualScreenWidth;
                                    else this.Width = tmp_d;
                                }
                                break;
                            case "height":
                                if (Double.TryParse(ertek, out tmp_d))
                                {
                                    if (tmp_d < this.MinHeight) this.Width = this.MinHeight;
                                    else if (tmp_d > SystemParameters.VirtualScreenHeight) this.Height = SystemParameters.VirtualScreenHeight;
                                    else this.Height = tmp_d;
                                }
                                break;
                            case "windowstate":
                                if (Enum.TryParse(ertek, out WindowState tmp_ws)) this.WindowState = tmp_ws;
                                break;
                            case "file":
                                if (ertek != null)
                                {
                                    this.textBox_file.Text = ertek;
                                    this._File = ertek;

                                    Thread t = new Thread(delegate ()
                                    {
                                        Action a = () =>
                                        {
                                            this.betoltesAlatt = true;
                                            this.SafeOpenFile();
                                            this.betoltesAlatt = false;
                                        };
                                        Dispatcher.BeginInvoke(a);
                                    });
                                    t.Start();
                                }
                                break;
                            case "oszlop":
                                if (int.TryParse(ertek, out int tmp_i) && tmp_i > -1) this._PreferaltOszlop = tmp_i;
                                break;
                            case "email":
                                if (ertek != null) this.textBox_email.Text = ertek.Replace(";", "\r\n");
                                break;
                            case "lejart_ev":
                                if (ertek != null) this.textBox_lejart_ev.Text = ertek;
                                break;
                            case "lejart_honap":
                                if (ertek != null) this.textBox_lejart_honap.Text = ertek;
                                break;
                            case "lejart_nap":
                                if (ertek != null) this.textBox_lejart_nap.Text = ertek;
                                break;
                            case "figyelmeztet_ev":
                                if (ertek != null) this.textBox_figyelmeztet_ev.Text = ertek;
                                break;
                            case "figyelmeztet_honap":
                                if (ertek != null) this.textBox_figyelmeztet_honap.Text = ertek;
                                break;
                            case "figyelmeztet_nap":
                                if (ertek != null) this.textBox_figyelmeztet_nap.Text = ertek;
                                break;
                            case "inditas_windowssal":
                                if (Boolean.TryParse(ertek, out bool tmp_b)) this.checkBox.IsChecked = tmp_b;
                                break;
                        }
                    }
                }
            }
            this.betoltesAlatt = false;
        }

        private void button_email_Click(object sender, RoutedEventArgs e)
        {
            this.button.IsEnabled = false;

            EmailKuldes();

            this.button.IsEnabled = true;
        }


        private void EmailKuldes()
        {
            using (SmtpClient smtp = new SmtpClient())
            using (MailMessage mail = new MailMessage())
            {
                // címbeállítások:
                mail.From = new MailAddress("asszistantbugreporter@gmail.com");

                foreach (string item in this.textBox_email.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mail.To.Add(new MailAddress(item));
                }

                mail.Subject = "Lejáró műszakik emlékeztető - " + DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss");
                StringBuilder sb = new StringBuilder();

                if (this.comboBox_oszlop.SelectedIndex > -1)
                {
                    sb.Append("<p>Az alábbi műszakik fognak hamarosan lejárni:</p>");
                    sb.Append("<table>");
                    bool elso = true;
                    foreach (DataRow item in this.dt.Rows)
                    {
                        if (elso)
                        {
                            sb.Append("<thead>");
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                sb.Append("<th>" + item[i] + "</th>");
                            }
                            sb.Append("</thead>");
                            elso = false;
                            sb.Append("<tbody>");
                        }

                        else if (DateTime.TryParseExact(item[this.comboBox_oszlop.SelectedIndex].ToString(), "yyyy.MM.dd", null, DateTimeStyles.None, out DateTime tmp))
                        {
                            if (this.IsLejart(this.Lejarat(tmp)) > 0)
                            {
                                sb.Append("<tr>");
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    sb.Append("<td>" + item[i] + "</td>");
                                }
                                sb.Append("</tr>");
                            }
                        }
                    }
                    sb.Append("</tbody>");
                    sb.Append("</table>");
                }

                mail.IsBodyHtml = true;
                mail.Body = sb.ToString();

                // smtp beállítás:
                smtp.Host = "smtp.gmail.com";
                smtp.Port = 587;
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential("asszistantbugreporter@gmail.com", "19920730Hv9lvv");

                //email küldése:
                try { smtp.Send(mail); }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Hiba a levél küldésekor!", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            switch (this.WindowState)
            {
                case WindowState.Normal:
                    this.ShowInTaskbar = true;
                    break;
                case WindowState.Minimized:
                    this.ShowInTaskbar = false;
                    break;
            }
            this.BealitastMent();
        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
            this.SetStartup(true);
            this.BealitastMent();
        }

        private void checkBox_Unchecked(object sender, RoutedEventArgs e)
        {
            this.SetStartup(false);
            this.BealitastMent();
        }

        private void Window_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                Process.Start("http://users.atw.hu/csernay/Muszaki_jelzo/Sugo/muszaki_emlekezteto.html");
            }
            if (e.Key == Key.Escape)
            {
                this.WindowState = WindowState.Minimized;
                this.BealitastMent();
            }
        }

        public void SetStartup(bool allapot)
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (allapot) rk.SetValue("ExcelEmlekezteto", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" -minimized");
            else rk.DeleteValue("ExcelEmlekezteto", false);
        }

        [System.Runtime.InteropServices.DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
    }
}
