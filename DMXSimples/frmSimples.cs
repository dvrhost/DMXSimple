using System;
using System.CodeDom;
using System.ComponentModel;
using System.Linq;
using System.Timers;
using System.Windows.Forms;
using Tatais.DMX.Properties;

namespace Tatais.DMX
{
    public partial class frmSimples : Form
    {
        public bool TimerEnabled = false;
        private static System.Timers.Timer aTimer;
        private static int ListElementNum;
        

        static Automation automation = new Automation();

        private bool SelectedChan1, SelectedChan2, SelectedChan3, SelectedChan4, SelectedChan5,
                SelectedChan6, SelectedChan7, SelectedChan8;
        private bool SelectedRed, SelectedGreen, SelectedBlue, SelectedWhite;
        private bool SelectedZoom;

        private bool SelectedChan1_Preset_1, SelectedChan2_Preset_1, SelectedChan3_Preset_1,
                SelectedChan4_Preset_1, SelectedChan5_Preset_1,
                SelectedChan6_Preset_1, SelectedChan7_Preset_1, SelectedChan8_Preset_1;
        private bool SelectedRed_Preset_1, SelectedGreen_Preset_1, SelectedBlue_Preset_1, 
                SelectedWhite_Preset_1, SelectedZoom_Preset_1;

        private bool SelectedChan1_Preset_2, SelectedChan2_Preset_2, SelectedChan3_Preset_2,
                SelectedChan4_Preset_2, SelectedChan5_Preset_2,
                SelectedChan6_Preset_2, SelectedChan7_Preset_2, SelectedChan8_Preset_2;
        private bool SelectedRed_Preset_2, SelectedGreen_Preset_2, SelectedBlue_Preset_2,
                SelectedWhite_Preset_2, SelectedZoom_Preset_2;

        
        private void Save_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Application.StartupPath;
                openFileDialog.Filter = "XML Files (*.xml)|*.xml";
                if(openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filename = openFileDialog.FileName;
                    automation.LoadData(filename);
                }
            }
            //automation.SaveData("C:\\Install\\DMXdata.xml");
            

        }
        public void StartShow()
        {
            ListElementNum = 0;
            if (automation.automationListDMXData != null)
            {
                stopProgram.Enabled = !stopProgram.Enabled;
                startProgram.Enabled = !startProgram.Enabled;
                automation.automationListDMXData.Sort();
                if (automation.automationListDMXData.Count > 0)
                {
                    SetTimer(automation.automationListDMXData[ListElementNum].Sec);
                }
            }
        }
        public void StopShow()
        {
            if(aTimer!=null )
                aTimer.Stop();
            stopProgram.Enabled = !stopProgram.Enabled;
            startProgram.Enabled = !startProgram.Enabled;
        }
        private void SetTimer(int sTime)
        {
            if (aTimer == null)
            {
                aTimer = new System.Timers.Timer(sTime * 1000);
                aTimer.Elapsed += aTimer_Elapsed;
                aTimer.Enabled = true;
                TimerEnabled = true;
            }
            else
            {
                aTimer.Interval= sTime*1000;
                aTimer.Enabled = true;
            }
        }

        private void aTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            aTimer.Stop();
            TimerEnabled = false;
            if (ListElementNum >= 0 && ListElementNum < automation.automationListDMXData.Count)
            {
                var unit = automation.automationListDMXData[ListElementNum].Unit;
                var red = automation.automationListDMXData[ListElementNum].Red;
                var green = automation.automationListDMXData[ListElementNum].Green;
                var blue = automation.automationListDMXData[ListElementNum].Blue;
                var white = automation.automationListDMXData[ListElementNum].White;
                var zoom = automation.automationListDMXData[ListElementNum].Zoom;
                ListElementNum++;
                SetUnitsData(unit, red, green, blue, white, zoom);
                if (ListElementNum < automation.automationListDMXData.Count)
                    SetTimer(automation.automationListDMXData[ListElementNum].Sec);
                else
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        StopShow();
                        MessageBox.Show("Automation program compleated", "Automation") ;
                    }));
                    
                }
            }
        }
        /// <summary>
        /// Set data to port DMX512
        /// </summary>
        /// <param name="unit">Unit number</param>
        /// <param name="red">red value</param>
        /// <param name="green">green value</param>
        /// <param name="blue">blue value</param>
        /// <param name="white">white value</param>
        /// <param name="zoom">zoom value</param>
        private void SetUnitsData(int unit, int red, int green, int blue, int white, int zoom)
        {
            int ChanelR;
            int ChanelG;
            int ChanelB;
            int ChanelW;
            int ChanelZ;
            switch (unit)
            {
                case 1:
                    ChanelR = 1;
                    ChanelG = 2;
                    ChanelB = 3;
                    ChanelW = 4;
                    ChanelZ = 5;
                    break;
                case 2:
                    ChanelR = 6;
                    ChanelG = 7;
                    ChanelB = 8;
                    ChanelW = 9;
                    ChanelZ = 10;
                    break;
                case 3:
                    ChanelR = 11;
                    ChanelG = 12;
                    ChanelB = 13;
                    ChanelW = 14;
                    ChanelZ = 15;                    
                    break;
                case 4:
                    ChanelR = 16;
                    ChanelG = 17;
                    ChanelB = 18;
                    ChanelW = 19;
                    ChanelZ = 20;
                    break;
                case 5:
                    ChanelR = 21;
                    ChanelG = 22;
                    ChanelB = 23;
                    ChanelW = 24;
                    ChanelZ = 25;
                    break;
                case 6:
                    ChanelR = 26;
                    ChanelG = 27;
                    ChanelB = 28;
                    ChanelW = 29;
                    ChanelZ = 30;
                    break;
                case 7:
                    ChanelR = 31;
                    ChanelG = 32;
                    ChanelB = 33;
                    ChanelW = 34;
                    ChanelZ = 35;
                    break;
                case 8:
                    ChanelR = 36;
                    ChanelG = 37;
                    ChanelB = 38;
                    ChanelW = 39;
                    ChanelZ = 40;
                    break;
                default:
                    ChanelR = 0;
                    ChanelG = 0;
                    ChanelB = 0;
                    ChanelW = 0;
                    ChanelZ = 0;
                    break;
            }
            buffer[ChanelR - 1] = Convert.ToByte(red);
            buffer[ChanelG - 1] = Convert.ToByte(green);
            buffer[ChanelB - 1] = Convert.ToByte(blue);
            buffer[ChanelW - 1] = Convert.ToByte(white);
            buffer[ChanelZ - 1] = Convert.ToByte(zoom);

            if (dmxCommunicator != null)
            {
                dmxCommunicator.SetBytes(buffer);
            }
        }

        private void LoadData_Click(object sender, EventArgs e)
        {
            StartShow();
        }

        private void stopProgram_Click(object sender, EventArgs e)
        {
            StopShow();
        }

        private int MainValue_Preset_1, MainValue_Preset_2, MainTrackBarValue;

        private bool PresetSelected_1 = true;
        private bool PresetSelected_2 = false;
        private int[] Preset_1_Data;
        private int[] Preset_2_Data;
        
        

        ColorDialog ColorPicup = new ColorDialog();
        private byte[] buffer = new byte[512];
        private DMXCommunicator dmxCommunicator = null;
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //load data
            PresetStore();
            PresetSelected_2 = true;
            PresetSelected_1 = false;
            PresetLoad();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //load data
            PresetStore();
            PresetSelected_2 = false;
            PresetSelected_1 = true;
            PresetLoad();
        }  
            /// <summary>
            /// Load preset setting for channels.
            /// </summary>
        private void PresetLoad()
        {
            
            
            if (PresetSelected_1)
            {
                //Load settings for preset 1
                SelectedGreen = SelectedGreen_Preset_1;
                SelectedBlue = SelectedBlue_Preset_1;
                SelectedRed = SelectedRed_Preset_1;
                SelectedWhite = SelectedWhite_Preset_1;
                SelectedZoom = SelectedZoom_Preset_1;
                SelectedChan1 = SelectedChan1_Preset_1;
                SelectedChan2 = SelectedChan2_Preset_1;
                SelectedChan3 = SelectedChan3_Preset_1;
                SelectedChan4 = SelectedChan4_Preset_1;
                SelectedChan5 = SelectedChan5_Preset_1;
                SelectedChan6 = SelectedChan6_Preset_1;
                SelectedChan7 = SelectedChan7_Preset_1;
                SelectedChan8 = SelectedChan8_Preset_1;
                MainTrackBarValue = MainValue_Preset_1;
                if (Preset_1_Data != null)
                {
                    int i = 1;
                    foreach (int element in Preset_1_Data)
                    {
                        var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i), true).FirstOrDefault() as NumericUpDown;
                        if (numericUpDown != null)
                        { numericUpDown.Value = element; }
                        i++;
                    }
                }
            }
            else if(PresetSelected_2)
            {
                //Load settings for preset 2
                SelectedGreen = SelectedGreen_Preset_2;
                SelectedBlue = SelectedBlue_Preset_2;
                SelectedRed = SelectedRed_Preset_2;
                SelectedWhite = SelectedWhite_Preset_2;
                SelectedZoom = SelectedZoom_Preset_2;   
                SelectedChan1 = SelectedChan1_Preset_2;
                SelectedChan2 = SelectedChan2_Preset_2;
                SelectedChan3 = SelectedChan3_Preset_2;
                SelectedChan4 = SelectedChan4_Preset_2;
                SelectedChan5 = SelectedChan5_Preset_2;
                SelectedChan6 = SelectedChan6_Preset_2;
                SelectedChan7 = SelectedChan7_Preset_2;
                SelectedChan8 = SelectedChan8_Preset_2;
                MainTrackBarValue = MainValue_Preset_2;
                //Load saved settings from Properties.Settings for Preset 2 and set it on GUI elements
                if (Preset_2_Data != null)
                {
                    int i = 1;
                    foreach (int element in Preset_2_Data)
                    {
                        var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i), true).FirstOrDefault() as NumericUpDown;
                        if (numericUpDown != null)
                        { numericUpDown.Value = element; }
                        i++;
                    }
                }
            }
            //Update GUI
            for (int chanel = 1; chanel <= 8; chanel++)
            {
                var checkboxunit = this.Controls.Find(string.Format("groupUnit{0}", chanel), true).FirstOrDefault() as CheckBox;
                bool value;
                switch(chanel)
                {
                    case 1:
                        value = SelectedChan1;
                        break;
                        case 2:
                        value = SelectedChan2;
                        break;
                        case 3:
                        value = SelectedChan3;
                        break;
                        case 4:
                        value = SelectedChan4;
                        break;
                        case 5:
                        value = SelectedChan5;
                        break;
                        case 6:
                        value = SelectedChan6;
                        break;
                        case 7:
                        value = SelectedChan7;
                        break;
                        case 8:
                        value = SelectedChan8;
                        break;
                        default: value = false; break;
                }
                if (checkboxunit != null)
                    checkboxunit.Checked = value;
            }
            var checkbox = this.Controls.Find("checkRed", true).FirstOrDefault() as CheckBox;
            if (checkbox != null)
                checkbox.Checked = SelectedRed;
            checkbox = this.Controls.Find("checkGreen", true).FirstOrDefault() as CheckBox;
            if (checkbox != null)
                checkbox.Checked = SelectedGreen;
            checkbox = this.Controls.Find("checkBlue", true).FirstOrDefault() as CheckBox;
            if (checkbox != null)
                checkbox.Checked = SelectedBlue;
            checkbox = this.Controls.Find("checkWhite", true).FirstOrDefault() as CheckBox;
            if (checkbox != null)
                checkbox.Checked = SelectedWhite;
            checkbox = this.Controls.Find("checkZoom", true).FirstOrDefault() as CheckBox;
            if (checkbox != null)
                checkbox.Checked = SelectedZoom;
            var numvalue = this.Controls.Find("numericUpDownMaster", true).FirstOrDefault() as NumericUpDown;
            if(numvalue !=null)
                numvalue.Value = MainTrackBarValue;
        }
        /// <summary>
        /// Save current settings of selected preset.
        /// </summary>
        private void PresetStore()
        {
            
            if (PresetSelected_1)
            {
                SelectedWhite_Preset_1 = SelectedWhite;
                SelectedRed_Preset_1 = SelectedRed;
                SelectedGreen_Preset_1 = SelectedGreen;
                SelectedBlue_Preset_1 = SelectedBlue;
                SelectedZoom_Preset_1 = SelectedZoom;
                SelectedChan1_Preset_1 = SelectedChan1;
                SelectedChan2_Preset_1 = SelectedChan2;
                SelectedChan3_Preset_1 = SelectedChan3;
                SelectedChan4_Preset_1 = SelectedChan4;
                SelectedChan5_Preset_1 = SelectedChan5;
                SelectedChan6_Preset_1 = SelectedChan6;
                SelectedChan7_Preset_1 = SelectedChan7;
                SelectedChan8_Preset_1 = SelectedChan8;
                var trackbar = this.Controls.Find("numericUpDownMaster", true).FirstOrDefault() as NumericUpDown;
                if(trackbar != null)
                    MainValue_Preset_1 = Convert.ToInt32(trackbar.Value);
                Preset_1_Data = new int[40];
                for (int i = 0; i < 40; i++)
                {
                    var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i + 1), true).FirstOrDefault() as NumericUpDown;

                    if (numericUpDown != null)
                        Preset_1_Data[i] = Convert.ToInt32(numericUpDown.Value);
                }
            }
            else if (PresetSelected_2)
            {
                SelectedWhite_Preset_2 = SelectedWhite;
                SelectedRed_Preset_2 = SelectedRed;
                SelectedGreen_Preset_2 = SelectedGreen;
                SelectedBlue_Preset_2 = SelectedBlue;
                SelectedZoom_Preset_2 = SelectedZoom;
                SelectedChan1_Preset_2 = SelectedChan1;
                SelectedChan2_Preset_2 = SelectedChan2;
                SelectedChan3_Preset_2 = SelectedChan3;
                SelectedChan4_Preset_2 = SelectedChan4;
                SelectedChan5_Preset_2 = SelectedChan5;
                SelectedChan6_Preset_2 = SelectedChan6;
                SelectedChan7_Preset_2 = SelectedChan7;
                SelectedChan8_Preset_2 = SelectedChan8;
                var trackbar = this.Controls.Find("numericUpDownMaster", true).FirstOrDefault() as NumericUpDown;
                if(trackbar !=null)
                    MainValue_Preset_2 = Convert.ToInt32(trackbar.Value);
                //Save channel values to array for future save in Properties.Settings at App shutdown.
                Preset_2_Data = new int[40];
                for (int i = 0; i < 40; i++)
                {
                    var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i + 1), true).FirstOrDefault() as NumericUpDown;

                    if (numericUpDown != null)
                        Preset_2_Data[i] = Convert.ToInt32(numericUpDown.Value);
                }
            }
        }
        private void checkRed_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = sender as CheckBox;
            SelectedRed =  checkbox.Checked;
            
        }

        private void checkGreen_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = sender as CheckBox;
            SelectedGreen = checkbox.Checked;
            
        }

        private void checkBlue_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = sender as CheckBox;
            SelectedBlue = checkbox.Checked;
            
        }

        private void checkWhite_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = sender as CheckBox;
            SelectedWhite = checkbox.Checked;
            
        }

        private void checkZoom_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = sender as CheckBox;
            SelectedZoom = checkbox.Checked;
            
        }

        private void trackBarMaster_Scroll(object sender, EventArgs e)
        {
            var trackBar = sender as TrackBar;
            var numericUpDown = this.Controls.Find("numericUpDownMaster", true).FirstOrDefault() as NumericUpDown;
            if (numericUpDown != null)
                numericUpDown.Value = trackBar.Value;
        }

        private void numericUpDownMaster_ValueChanged(object sender, EventArgs e)
        {
            var numericUpDown = sender as NumericUpDown;
            var trackbar = this.Controls.Find("trackBarMaster", true).FirstOrDefault() as TrackBar;
            
            if (trackbar != null)
                trackbar.Value = Convert.ToInt32(numericUpDown.Value);

            //Group control 
           
            if (SelectedChan1)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 1);
                if(SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 2);
                if(SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 3);
                if(SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 4);
                if(SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 5);
            }
            if (SelectedChan2)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 6);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 7);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 8);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 9);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 10);
            }
            if (SelectedChan3)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 11);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 12);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 13);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 14);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 15);
            }
            if (SelectedChan4)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 16);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 17);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 18);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 19);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 20);
            }
            if (SelectedChan5)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 21);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 22);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 23);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 24);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 25);
            }
            if (SelectedChan6)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 26);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 27);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 28);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 29);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 30);
            }
            if (SelectedChan7)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 31);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 32);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 33);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 34);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 35);
            }
            if (SelectedChan8)
            {
                if (SelectedRed)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 36);
                if (SelectedGreen)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 37);
                if (SelectedBlue)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 38);
                if (SelectedWhite)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 39);
                if (SelectedZoom)
                    SetChanelAt(Convert.ToInt32(numericUpDown.Value), 40);
            }
        }
        private void SetChanelAt(int value, int chanel)
        {
            var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", chanel), true).FirstOrDefault() as NumericUpDown;

            if (numericUpDown != null)
                numericUpDown.Value = value;
        }
        /// <summary>
        /// Open Color Picup dialog for set Led lights color.
        /// </summary>
        private void MasterColor_Click(object sender, EventArgs e)
        {
            //Group set color
            ColorPicup.AllowFullOpen = true;

            if (ColorPicup.ShowDialog() == DialogResult.OK)
            {
                var red = ColorPicup.Color.R;
                var green = ColorPicup.Color.G;
                var blue = ColorPicup.Color.B;
                if(SelectedChan1)
                    SetColors(red, green, blue, 1);
                if(SelectedChan2)
                    SetColors(red, green, blue, 2);
                if(SelectedChan3)
                    SetColors(red, green, blue, 3);
                if(SelectedChan4)
                    SetColors(red, green, blue, 4);
                if(SelectedChan5)
                    SetColors(red, green, blue, 5);
                if(SelectedChan6)
                    SetColors(red, green, blue, 6);
                if(SelectedChan7)
                    SetColors(red, green, blue, 7);
                if(SelectedChan8)
                    SetColors(red, green, blue, 8);
            }
        }

        

        public frmSimples()
        {
            InitializeComponent();
            this.Icon = Tatais.DMX.Properties.Resources.Icon;
            var portsList = DMXCommunicator.GetValidSerialPorts();
            cbxSerialPort.DataSource = new BindingSource(portsList, null);
            //Load saved data from Properties.Settings
            var Preset_1 = (string)Settings.Default["Preset_1_Data"];
            var Preset_2 = (string)Settings.Default["Preset_2_Data"];

            try
            {
                Preset_1_Data = Preset_1.Split(',').Select(x => Int32.Parse(x)).ToArray();
                
            }
            catch
            {
                Exception ex = new Exception("Wrong saved settings in Preset 1");
            }
            try
            {
                Preset_2_Data = Preset_2.Split(',').Select(x => Int32.Parse(x)).ToArray();
                
            }
            catch
            {
                Exception ex = new Exception("Wrong saved settings in Preset 2");
            }
            RestoreSettings();
           

        }
        protected override void OnClosing(CancelEventArgs e)
        {
            //Save actual presets data
            if (PresetSelected_1)
            {
                Preset_1_Data = new int[40];
                for (int i = 0; i < 40; i++)
                {
                    var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i + 1), true).FirstOrDefault() as NumericUpDown;

                    if (numericUpDown != null)
                        Preset_1_Data[i] = Convert.ToInt32(numericUpDown.Value);
                }
            }
            else if (PresetSelected_2)
            {
                Preset_2_Data = new int[40];
                for (int i = 0; i < 40; i++)
                {
                    var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i + 1), true).FirstOrDefault() as NumericUpDown;

                    if (numericUpDown != null)
                        Preset_2_Data[i] = Convert.ToInt32(numericUpDown.Value);
                }
            }
            //Save data to arrays
            if (Preset_1_Data != null)
            {
                string value1 = String.Join(",", Preset_1_Data.Select(x => x.ToString()));
                Properties.Settings.Default.Preset_1_Data = value1;
            }
            if (Preset_2_Data != null)
            {
                string value2 = String.Join(",", Preset_2_Data.Select(x => x.ToString()));
                Properties.Settings.Default.Preset_2_Data = value2;
            }
            //write data to Properties
            Properties.Settings.Default.Save();
            base.OnClosing(e);
        }
        /// <summary>
        /// Restore settings on App open for Preset 1
        /// </summary>
        private void RestoreSettings()
        {
            if (Preset_1_Data != null)
            {
                int i = 1;
                foreach (int element in Preset_1_Data)
                {
                    var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", i), true).FirstOrDefault() as NumericUpDown;
                    if (numericUpDown != null)
                    { numericUpDown.Value = element; }
                    i++;
                }
            }
        }
        private void TrackBar_Scroll(object sender, EventArgs e)
        {
            var trackBar = sender as TrackBar;
            var position = Convert.ToInt16(trackBar.Name.Substring(8));
            var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", position), true).FirstOrDefault() as NumericUpDown;

            if (numericUpDown != null)
                numericUpDown.Value = trackBar.Value;
        }

        private void NumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            var numericUpDown = sender as NumericUpDown;
            var position = Convert.ToInt16(numericUpDown.Name.Substring(13));
            var trackbar = this.Controls.Find(string.Format("trackbar{0}", position), true).FirstOrDefault() as TrackBar;

            if (trackbar != null)
                trackbar.Value = Convert.ToInt32(numericUpDown.Value);

            buffer[position-1] = Convert.ToByte(numericUpDown.Value);
            if (dmxCommunicator != null)
                dmxCommunicator.SetByte(position - 1, Convert.ToByte(numericUpDown.Value));
        }
        //Group color setup
        private void btn_ColorPickUp(object sender, EventArgs e)
        {
            var ColorPicupChanel = sender as Button;
            var position = Convert.ToInt16(ColorPicupChanel.Name.Substring(8));
            
            ColorPicup.AllowFullOpen = true;
           
            if (ColorPicup.ShowDialog() == DialogResult.OK)
            {
                var red = ColorPicup.Color.R;
                var green = ColorPicup.Color.G;
                var blue = ColorPicup.Color.B;
                SetColors( red, green, blue,position);
            }
        }
        /// <summary>
        /// Set color for sellected RGB Led Unit
        /// </summary>
        /// <param name="red">Red channel from 0 to 255</param>
        /// <param name="green">Green channel from 0 to 255</param>
        /// <param name="blue">Blue channel from 0 to 255</param>
        /// <param name="chanel">RGBLed Unit DMX512 num from 1 to 8</param>
        private void SetColors(byte red,byte green,byte blue,int chanel)
        {
            short ChanelR;
            short ChanelG;
            short ChanelB;
            switch (chanel)
            {
                case 1:
                    ChanelR = 1;
                    ChanelG = 2;
                    ChanelB = 3;
                    break;
                case 2:
                    ChanelR = 6;
                    ChanelG = 7;
                    ChanelB = 8;
                    break;
                case 3:
                    ChanelR = 11;
                    ChanelG = 12;
                    ChanelB = 13;
                    break;
                case 4:
                    ChanelR = 16;
                    ChanelG = 17;
                    ChanelB = 18;
                    break;
                case 5:
                    ChanelR = 21;
                    ChanelG = 22;
                    ChanelB = 23;
                    break;
                case 6:
                    ChanelR = 26;
                    ChanelG = 27;
                    ChanelB = 28;
                    break;
                case 7:
                    ChanelR = 31;
                    ChanelG = 32;
                    ChanelB = 33;
                    break;
                case 8:
                    ChanelR = 36;
                    ChanelG = 37;
                    ChanelB = 38;
                    break;
                default:
                    ChanelR = 0;
                    ChanelG = 0;
                    ChanelB = 0;
                    break;
            }
            if (ChanelR > 0)
            { var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", ChanelR), true).FirstOrDefault() as NumericUpDown;

                if (numericUpDown != null)
                    numericUpDown.Value = red;
            }
            if (ChanelB > 0)
            {
                var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", ChanelB), true).FirstOrDefault() as NumericUpDown;

                if (numericUpDown != null)
                    numericUpDown.Value = blue;
            }
            if (ChanelG > 0)
            {
                var numericUpDown = this.Controls.Find(string.Format("numericUpDown{0}", ChanelG), true).FirstOrDefault() as NumericUpDown;

                if (numericUpDown != null)
                    numericUpDown.Value = green;
            }
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (cbxSerialPort.Items.Count == 0)
                return;
            
            dmxCommunicator = new DMXCommunicator(cbxSerialPort.SelectedValue.ToString());
            dmxCommunicator.SetBytes(buffer);
            dmxCommunicator.Start();

            cbxSerialPort.Enabled = !cbxSerialPort.Enabled;
            btnStart.Enabled = !btnStart.Enabled;
            btnStop.Enabled = !btnStop.Enabled;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (dmxCommunicator != null && dmxCommunicator.IsActive)
            {
                dmxCommunicator.Stop();
                dmxCommunicator = null;
            }

            cbxSerialPort.Enabled = !cbxSerialPort.Enabled;
            btnStart.Enabled = !btnStart.Enabled;
            btnStop.Enabled = !btnStop.Enabled;
        }

        private void frmSimple_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dmxCommunicator != null && dmxCommunicator.IsActive)
            {
                dmxCommunicator.Stop();
                dmxCommunicator = null;
            }
        }

        private void groupUnit_CheckedChanged(object sender, EventArgs e)
        {
            var checkboxselected = sender as CheckBox;
            var position = Convert.ToInt16(checkboxselected.Name.Substring(9));
            switch (position)
            { 
                case 1:
                    SelectedChan1 = checkboxselected.Checked;
                    break;
                case 2:
                    SelectedChan2 = checkboxselected.Checked;
                    break;
                case 3:
                    SelectedChan3 = checkboxselected.Checked;
                    break;
                case 4:
                    SelectedChan4 = checkboxselected.Checked;
                    break;
                case 5:
                    SelectedChan5 = checkboxselected.Checked;
                    break;
                case 6:
                    SelectedChan6 = checkboxselected.Checked;
                    break;  
                case 7:
                    SelectedChan7 = checkboxselected.Checked;
                    break;
                case 8:
                    SelectedChan8 = checkboxselected.Checked;
                    break;
                default:
                    break;
            }
            
        }
        
    }
    
}
