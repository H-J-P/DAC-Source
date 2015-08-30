
#region Lizenz
/*
        ----------------------------------------------------------------------
         Copyright (c) 2014 Heinz-Joerg Puhlmann

         Hiermit wird unentgeltlich, jeder Person, die eine Kopie der Software
         und der zugehörigen Dokumentationen (die "Software") erhält, die
         Erlaubnis erteilt, uneingeschränkt zu benutzen, inklusive und ohne
         Ausnahme, dem Recht, sie zu verwenden, kopieren, ändern, fusionieren,
         verlegen, verbreiten, unter-lizenzieren und/oder zu verkaufen, und 
         Personen, die diese Software erhalten, diese Rechte zu geben, unter
         den folgenden Bedingungen:

         Der obige Urheberrechtsvermerk und dieser Erlaubnisvermerk sind in 
         alle Kopien oder Teilkopien der Software beizulegen.

         DIE SOFTWARE WIRD OHNE JEDE AUSDRÜCKLICHE ODER IMPLIZIERTE GARANTIE 
         BEREITGESTELLT, EINSCHLIESSLICH DER GARANTIE ZUR BENUTZUNG FÜR DEN
         VORGESEHENEN ODER EINEM BESTIMMTEN ZWECK SOWIE JEGLICHER 
         RECHTSVERLETZUNG, JEDOCH NICHT DARAUF BESCHRÄNKT. IN KEINEM FALL SIND
         DIE AUTOREN ODER COPYRIGHTINHABER FÜR JEGLICHEN SCHADEN ODER SONSTIGE
         ANSPRUCH HAFTBAR ZU MACHEN, OB INFOLGE DER ERFÜLLUNG VON EINEM 
         VERTRAG, EINEM DELIKT ODER ANDERS IM ZUSAMMENHANG MIT DER BENUTZUNG
         ODER SONSTIGE VERWENDUNG DER SOFTWARE ENTSTANDEN
        ----------------------------------------------------------------------
*/
#endregion

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using SimpleSolutions.Usb;
using WindowsInput;

namespace DAC
{
    /*
		---------------------------------------------------------------------------------------------------------------------
		D.A.C. DCS Arcaze Communicator
		---------------------------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------------------------------
		Version 0.900 Beta  DLL V6                                                                                 06.05.2014
		---------------------------------------------------------------------------------------------------------------------
		0.900    Versionierung eingeführt                                                                          06.05.2014
		0.901    Use Helios format; ValueGrabber() changed; database changed.                                      07.05.2014
		0.902    Database relations changed.                                                                       08.05.2014
		0.903    More then one Dot in a string                                                                     09.05.2014
		0.904    Added Modultyps to the database.
		0.904    Added resolution to LED table for LED - driver 3 (dimming)
		0.905    Database fields cleanup; default values added;                                                    10.05.2014
		0.906    Added ModuleType and button Init to tab Test;                                                     11.05.2014
		0.906    Added Resolution to Test tab; Added Value to Test tab
		0.907    Added Init for LED-Driver 2 and 3;                                                                12.05.2014
		         Database table "LEDs" field "Active" changed to bool;
		         Database table "config" all fields changed -> AllowDBNull = false; DefaultValue = "";
		---------------------------------------------------------------------------------------------------------------------
		0.908    Using of DLL V5.5;                                                                                13.05.2014
		---------------------------------------------------------------------------------------------------------------------
		0.909    Added change event for "Connected Arcaze" ComboBox and activate this arcaze.                      15.05.2014
		0.910    Reset button visible; automated resetprozess available incl. init for all types of driverboards.  16.05.2014
		0.911    Bug fixes;                                                                                        17.05.2014
		0.912    timerIntervall = 20ms; => 50 fps to get all packages. Importend: The export.lua should use 25 fps 17.05.2014
		0.913    Added checkbox "log all actions"                                                                  18.05.2014
		         Added button [Stop]. All packages processing are stoped or started.
		0.914    Cleanup logs                                                                                      19.05.2014
		0.915    Added LED-Driver functions                                                                        20.05.2014
		0.916    Added CheckLengthOfDisplayValue(); something like this "1.7654321" is now possible.               21.05.2014
		0.917    LED-Driver 2 and 3 is working incl. dimming.
		0.918    Code cleanup; Bug fixes; Change log time format to "HH:mm:ss.fff";                                22.05.2014
		         Log expanded; Tab Masterdata -> Arcazes found in the past -> added ModulType. That shows the
		         first connected driver modul. 
		0.919    New section "Arcaze test" in tab "Test". That test all LED and Displays.                          23.05.2014
		0.920    Added SystemCheck on startup;                                                                     24.05.2014
		         First System optimising RAM and CPU usages;
		0.921    Activate Arcaze after seletion tab "Test";                                                        25.05.2014
		         "0-9 , ." are only valid character for parameters. 
		         Parameters with other character will be corrected.
		0.922    Switch/Load a new profile from data package ('File=A-10C:' parameter = selfData.Name from lua)    26.05.2014
		0.923    Tab 'Configuration': Added textbox for 'Interval time [ms]';                                      27.05.2014
		         Added checkbox for 'Write logs to HD'; The file name is 'log.txt'; 
		         Added textbox for 'Test data package'; Added 'log all actions' to config.
		         Tab 'Test': Added button 'Send the test data package from tab 'Configuration''
		0.924    Tab 'Displays' and 'LEDs' column 'init' checkbox added.                                           28.05.2014
		0.925    Bug fixes; Stability update;                                                                      01.06.2014
		         Added columns: dM == dimming mode, dT == dimming time, dV == dimming value
		0.926    Added automatic dimming for LED-driver 3                                                          02.06.2014
		0.927    Bug fixes; All display segments after 'init' are now off;                                         03.06.2014
		0.928    Added sliders for brightness; After load new config during gameplay, no 'init' any more;          04.06.2014
		0.929    Bug fixes;                                                                                        05.06.2014
		0.930    Optimizing/Speedup file read and 'init'                                                           06.06.2014
		0.931    Code optimizing;                                                                                  07.06.2014
		0.932    Optimize blinking.                                                                                08.06.2014
		0.933    Tab Experimental added. Keystrokes for application like VAC or TrackIR.                           09.06.2014
		0.934    Added Keystrokes for joysticks.                                                                   10.06.2014
		0.935    Bug fixes;                                                                                        15.06.2014
		0.936    Database optimising; Added tab 'Switches'; first tests and corrections                            17.06.2014
		0.937    Added table 'Devices' and 'Devices and Switches';                                                 18.06.2014
		0.938    Now Switches are ready to send data back to DCS;                                                  19.06.2014
		         Added 'Value Off' and 'Value On' to tab switches for rotary switches;
		0.939    Optimizings; now sent network package as broadcast                                                23.06.2014
		0.940    Bug fixes;                                                                                        24.06.2014
		0.941    Added tab 'Encoders'; Added tables 'Encoder' and 'MultipostionSwitch';                            25.06.2014
		0.942    Bug fixes;                                                                                        26.06.2014
		0.943    Added tab 'Keystrokes'; Added table 'Keystrokes';                                                 27.06.2014
		0.944    Init and test function for encoders implemented; Code optimizing; Bug fixes;                      28.06.2014
		         Negative test of encoders. I get only '0000'. What is wrong?
		         Before the program finished, all LEDs and displays goes in the off mode;
		         Positive test of display driver 32. It's nice and working.
		0.945    After many restarts of DCS, I believe, we have now our first stable version with a great          29.06.2014
		         performance. I used for this tests our export.lua script.
		0.946    Added keystrokes for switches. It's now active. None progress with encoders.                      30.06.2014
		0.947    Added logic for encoders.                                                                         02.07.2014
		0.948    Change DTA=stop to DAC=stop;                                                                      03.07.2014
		0.949    Bug fixes; Database changed (Table DeviceType and Modul erased); Added 'I/O Changed';             06.07.2014
		         Added automatic database repairing.
		0.950    Hide tab 'Encoders';                                                                              08.07.2014
		0.951    Changed assembly info; Name changed from DTA.exe to DAC.exe; Save filter is now *.xml;            11.07.2014
		         Added checkbox 'Reverse' for displays. Tab 'Encoders' is again visible.
		
		---------------------------------------------------------------------------------------------------------------------
    */
    public partial class FormMain : Form
    {
        #region member

        enum State
        {
            sendKeys,
            readfile,
            startup,
            init,
            systemCheck,
            checkLEDS,
            checkDisplays,
            run,
            reset,
            stop
        }

        State timerstate = State.startup;
        private ArcazeHid arcazeHid = new ArcazeHid();

        private DataGridViewComboBoxCell comboCell = null;

        private List<DeviceInfo> allArcaze = new List<DeviceInfo>(16);
        private List<DeviceInfo> presentArcaze = new List<DeviceInfo>(16);
        private List<ArcazeDevice> arcazeDevice = new List<ArcazeDevice>();

        private List<byte> Digits = new List<byte>(8);

        private DataRow[] findArcaze = new DataRow[] { };
        private DataRow[] findDriver = new DataRow[] { };
        private DataRow[] rows = new DataRow[] { };
        private DataRow[] clickableRows = new DataRow[] { };

        private DataRow row = null;
        Boolean[] checkBoxDP = new Boolean[8] { false, false, false, false, false, false, false, false };
        Boolean[] checkBox = new Boolean[8];

        private Thread udpThread = null;

        private double dimmingInterval = 0.0;
        private double dimmingQuotient = 0.0;
        private double dimmingValue = 0.0;
        private double pinValue = 0;
        private double sendKeyDelay = 0;
        private double sendKeyValue = 0;

        private bool arcazeActive = false;
        private bool arcazeDeviceFound = false;
        private bool arcazeFound = false;
        private bool checkTest = false;
        private bool initDone = false;
        private bool ledOn = false;
        private bool loadFile = false;
        private bool lStateEnabled = true;
        private bool noValue = false;
        private const bool off = false;
        private const bool on = true;
        private bool reverse = false;
        private bool sendOpenKeysTimeBased = true;
        private bool stop = false;
        private bool switchesSendNotZero = false;
        private bool useLeftPadding = false;
        private bool isReverse = false;

        private byte characterValue = 0;

        private int arcazeDeviceIndex = 0;
        private int brightness = 0;
        private int buttonID = 0;
        private int channel = 0;
        private int clickable_ID = 0;
        private int dezValueByte = 0;
        private int displayCount = 0;
        private int deviceID = 0;

        private double encoderCalcValue = 0;
        private double encoderDentValue = 0;
        private double encoderMaxValue = 0;
        private double encoderMinValue = 0;

        private int encoderDeltaValue = 0;
        private int encoderIdentity = 0;

        private int encoderNewValue = 0;
        private int encoderOldValue = 0;
        private int encoderPinAvalue = 0;
        private int encoderPinBvalue = 0;
        private int encoderReturnValue = 0;

        private int encoderNo = 0;
        private string encoderName = "";
        private int mainTestLoop = 2;
        private int module = 0;
        //private int numberDigit = 0;
        private int numberInputChanged = 0;
        private int loopCounter = 0;
        private int loopMax = 50;
        private int pin = 0;
        private int pinVal = 0;
        private int port = 0;
        private int portValue = 0;
        //private int posDot = -1;
        private int posStart = 0;
        private int posEnd = 0;
        private int resolution = 1;
        private int segmentIndex = 0;
        private int testLoop = 0;
        private int timerInterval = 50;
        private int timerIntervalTest = 100;
        private int timerIntervalReset = 5000;
        private int type = 0;

        private uint resolutionValue = 1;

        private double encoderLastValue = 0;

        private string[] digit = new string[8] { "", "", "", "", "", "", "", "" };
        private string arcazeAddress = "";
        private string arcazeFromGrid = "";
        private string button = "";
        private string buttonsPressed = "";
        private string dataSetFilename = "\\Dataset.xml";
        private string dcsExportID = "";
        private string devAddress = "";
        private string dezValue = "";
        private string digitMask = "";
        private string digitString = "";
        private string fragment = "";
        private string fragmentID = "";
        private string gotData = "";
        private string maskHex = "";
        private string newline = System.Environment.NewLine;
        private string newValue = "";
        private string lastFile = "";
        private string logFile = Application.StartupPath + "\\" + "log.txt";
        private string pattern = "";
        private string moduleTyp = "";
        private string package = "";
        private string pinString = "";
        private string processName = "";
        private string readFile = "";
        private string receivedData = "";
        private string receivedDataBackup = "";
        private const string searchStringForFile = "File";
        private string sendKey = "";
        private string optionKeyCode = "";
        private string optionKeyCodeTwo = "";
        private string valueON = "1.0";
        private string valueOff = "0.0";

        #endregion

        public FormMain()
        {
            InitializeComponent();

            ImportExport.LogMessage("Application started ...", true);
            ImportExport.LogMessage(labelVersion.Text, true);
            ImportExport.XmlToDataSet(Application.StartupPath + "\\" + "config.xml", dataSetConfig);

            checkBoxWriteLogsToHD.Checked = Convert.ToBoolean(dataSetConfig.Tables["Config"].Rows[0]["WriteToHD"]);
            checkBoxLog.Checked = Convert.ToBoolean(dataSetConfig.Tables["Config"].Rows[0]["LogAllActions"]);
            textBoxLastFile.Text = dataSetConfig.Tables["Config"].Rows[0]["LastFile"].ToString();

            textBoxPortListener.Text = dataSetConfig.Tables["Config"].Rows[0]["PortListener"].ToString();
            textBoxPortSender.Text = dataSetConfig.Tables["Config"].Rows[0]["PortSender"].ToString();

            textBoxTestDataPackage.Text = dataSetConfig.Tables["Config"].Rows[0]["TestData"].ToString() + newline;

            rows = new DataRow[] { };
            rows = dataSetConfig.Tables["SendKeysTimeBased"].Select("Active=" + true + " AND Sended=" + true);

            for (int n = 0; n < rows.Length; n++)
            {
                rows[n]["Sended"] = false;
                rows[n]["Value"] = 0;
            }
            dataSetConfig.Tables["SendKeysTimeBased"].AcceptChanges();

            InitConfig();

            if (textBoxLastFile.Text.IndexOf(".") == -1)
                textBoxLastFile.Text += ".xml";

            ImportExport.XmlToDataSet(Application.StartupPath + "\\" + textBoxLastFile.Text, dataSetDisplaysLEDs);
            readFile = textBoxLastFile.Text;
            lastFile = readFile;

            textBoxIntervalTimer.Text = dataSetConfig.Tables["Config"].Rows[0]["IntervalTimer"].ToString();
            ImportExport.LogMessage("Interval timer: " + textBoxIntervalTimer.Text + "[ms]", true);

            if (File.Exists(logFile))
            {
                try { File.Delete(logFile); }
                catch (Exception f)
                {
                    ImportExport.LogMessage("File.Delete ... " + f.ToString(), true);
                }
            }
            InitTables();

            //tabControl1.TabPages.Remove(Info);
            tabControl1.TabPages.Remove(Experimental);
            //tabControl1.TabPages.Remove(Keystrokes);

            //ShoutHello();

            StartTimer();
        }

        private void StartTimer()
        {
            timerInterval = Convert.ToInt32(textBoxIntervalTimer.Text);
            timerMain.Interval = timerInterval;
            timerMain.Tick += new EventHandler(timerMain_Tick);
            timerMain.Start();
        }

        private void timerMain_Tick(object sender, EventArgs e)
        {
            #region send keys timebased

            if (lStateEnabled && !stop)
            {
                if (sendOpenKeysTimeBased)
                {
                    lStateEnabled = false;

                    rows = new DataRow[] { };
                    rows = dataSetConfig.Tables["SendKeysTimeBased"].Select("Active=" + true + " AND Sended=" + false + " AND State = 'On startup'");

                    SendKeystrokesTimeBased(ref rows);

                    lStateEnabled = true;
                }
            }
            #endregion

            #region start listener

            if (UDP.CloseListener && textBoxPortListener.Text.Trim().Length >= 4)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    if (udpThread != null)
                        udpThread.Abort();

                    udpThread = new Thread(new ThreadStart(StartListener));
                    udpThread.IsBackground = true;
                    udpThread.Start();

                    this.Text = this.Text + "  (Config. with " + textBoxLastFile.Text + ")";

                    timerstate = State.startup;
                    lStateEnabled = true;
                }
            }
            #endregion

            #region switch to new file

            if (timerstate == State.readfile)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;
                    stop = true;

                    if (readFile.Length > 0 && readFile != lastFile)
                    {
                        if (File.Exists(Application.StartupPath + "\\" + readFile + ".xml"))
                        {
                            SwitchAll(off);
                            SwitchInverseOff();

                            try
                            {
                                ImportExport.LogMessage("Read configuration file : " + readFile + ".xml", true);

                                dataSetDisplaysLEDs.Clear();
                                ImportExport.XmlToDataSet(Application.StartupPath + "\\" + readFile + ".xml", dataSetDisplaysLEDs);
                                ImportExport.LogMessage("Used configuration file : " + readFile + ".xml", true);

                                InitDrivers();

                                lastFile = readFile;

                                textBoxLastFile.Text = lastFile + ".xml";
                                this.Text = "D.A.C. - DCS Arcaze Communicator  (Config. with " + textBoxLastFile.Text + ")";
                            }
                            catch (Exception f)
                            {
                                ImportExport.LogMessage("Cannot use  file : " + readFile + ".xml" + " ... " + f.ToString(), true);
                            }
                        }
                        else
                            ImportExport.LogMessage("File not found: " + readFile + ".xml ... ", true);

                        timerMain.Interval = timerInterval;
                        receivedData = receivedDataBackup;

                        if (initDone)
                            timerstate = State.run;
                    }
                    stop = false;
                    lStateEnabled = true;
                }
            }
            #endregion

            #region startup

            if (timerstate == State.startup)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    labelPleaseWait.Text = "Startup ... ";
                    timerMain.Interval = timerInterval;

                    if (dataSetConfig.Tables["Joysticks"].Rows.Count == 0)
                        Joystick.FindDevice(FormMain.ActiveForm);

                    arcazeFound = FindAllArcaze();

                    if (!arcazeFound)
                    {
                        labelPleaseWait.Text = "No Arcaze found";
                        ImportExport.LogMessage("No Arcaze found --> next try in 15 sec. <--", true);
                        timerMain.Interval = timerIntervalReset * 3;
                    }
                    else
                    {
                        buttonPackage.Visible = true;
                        buttonSendTestPattern.Visible = true;
                        labelOnOff.Visible = true;

                        groupBoxDisplay.Visible = true;
                        groupBoxPinSet.Visible = true;
                        //groupBoxEncoder.Visible = true;

                        labelPleaseWait.Text = "";
                        timerMain.Interval = timerInterval;
                        timerstate = State.init;
                    }
                    lStateEnabled = true;
                }
            }
            #endregion

            #region stop

            if (timerstate == State.stop)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    SwitchAll(off);

                    ImportExport.LogMessage(labelVersion.Text, true);
                    timerstate = State.run;

                    lStateEnabled = true;
                }
            }
            #endregion

            #region init

            if (timerstate == State.init)
            {
                if (arcazeFound)
                {
                    if (lStateEnabled)
                    {
                        lStateEnabled = false;

                        labelPleaseWait.Text = "Init is running. Please wait ..";

                        InitDrivers();
                        SwitchAll(on);
                        SwitchAll(off);
                        initDone = true;

                        testLoop = 0;
                        mainTestLoop = 0;
                        timerstate = State.systemCheck;

                        lStateEnabled = true;
                    }
                }
            }
            #endregion

            #region check LEDs

            if (timerstate == State.checkLEDS)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    if (testLoop < dataGridViewLEDs.RowCount)
                    {
                        if (receivedData == "")
                        {
                            try
                            {
                                dcsExportID = (string)dataGridViewLEDs.Rows[testLoop].Cells["dcsExportIDLEDs"].Value;
                                dataGridViewLEDs.Rows[testLoop].Cells["dimmValue"].Value = (0).ToString();

                                labelPleaseWait.Text = "LED Test " + (ledOn ? "On" : "Off");
                                pattern = ":" + dcsExportID + "=" + (ledOn ? "1" : "0") + ":";

                                receivedData = pattern;
                                testLoop++;
                            }
                            catch (Exception f)
                            {
                                ImportExport.LogMessage("Timerstate check LEDs ... " + f.ToString(), true);
                                testLoop = 0;
                                timerstate = State.checkDisplays;
                            }
                        }
                    }
                    else
                    {
                        testLoop = 0;
                        timerstate = State.checkDisplays;
                    }
                    lStateEnabled = true;
                }
            }
            #endregion

            #region check Displays

            if (timerstate == State.checkDisplays)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    if (testLoop < dataGridViewDisplays.RowCount)
                    {
                        if (receivedData == "")
                        {
                            comboCell = (DataGridViewComboBoxCell)dataGridViewDisplays.Rows[testLoop].Cells["dcsExportIDDisplays"];

                            try
                            {
                                if (comboCell.Value != null)
                                {
                                    dcsExportID = comboCell.Value.ToString();

                                    pattern = ":" + dcsExportID + "=" + (ledOn ? "88888888" : "-") + ":";
                                    labelPleaseWait.Text = "Display Test " + (ledOn ? "On" : "Off");

                                    receivedData = pattern;
                                }
                                testLoop++;
                            }
                            catch (Exception f)
                            {
                                ImportExport.LogMessage("Timerstate check Displays ... " + f.ToString(), true);
                            }
                        }
                    }
                    else
                    {
                        labelPleaseWait.Text = "";
                        buttonSendTestPattern.Enabled = true;
                        timerstate = State.systemCheck;
                    }
                    lStateEnabled = true;
                }
            }
            #endregion

            #region system check

            if (timerstate == State.systemCheck)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    if (mainTestLoop > 0)
                    {
                        timerMain.Interval = timerIntervalTest;

                        if (mainTestLoop == 2)
                            labelPleaseWait.Text = "Systemcheck 'On'";
                        else
                            labelPleaseWait.Text = "Systemcheck 'Off'";

                        mainTestLoop--;
                        SystemCheck();
                    }
                    else
                    {
                        try
                        {
                            comboBoxDevAddress.SelectedIndex = 0;
                            comboBoxConnector.SelectedIndex = 0;
                            comboBoxPin.SelectedIndex = 0;
                            comboBoxModuleType.SelectedIndex = 3;
                            comboBoxModul.SelectedIndex = 0;
                            //comboBoxModul
                        }
                        catch (Exception f)
                        {
                            ImportExport.LogMessage("Timerstate system check ... " + f.ToString(), true);
                        }

                        labelPleaseWait.Text = "";
                        buttonInit.Visible = false;
                        timerMain.Interval = timerInterval;
                        timerstate = State.run;
                    }
                    lStateEnabled = true;
                }
            }
            #endregion

            #region reset

            if (timerstate == State.reset)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    timerMain.Interval = timerIntervalReset;
                    labelPleaseWait.Text = "Reset Arcaze ... ";

                    ComboBoxArcaze.Items.Clear();
                    ComboBoxArcaze.Text = "";

                    for (int n = 0; n < presentArcaze.Count; n++)
                    {
                        arcazeHid.Connect(presentArcaze[n].Path);
                        arcazeHid.Command.CmdReset();
                        ImportExport.LogMessage(presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ") reseted .. ", true);
                    }
                    arcazeHid.Disconnect();
                    arcazeAddress = "";
                    timerstate = State.startup;

                    lStateEnabled = true;
                }
            }
            #endregion

            #region dimming / blinking

            if (timerstate == State.run && !stop && lStateEnabled)
            {
                for (int n = 0; n < dataGridViewLEDs.Rows.Count; n++)
                {
                    comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["modulTypeID"];

                    if (comboCell.Value == null)
                        type = 4;
                    else
                        type = comboCell.Items.IndexOf(comboCell.Value) + 1;

                    if (type != 1 && Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["dimmMode"].Value) != 0)
                    {
                        if (double.Parse(dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture) > 0.5)
                        {
                            if (dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value != null)
                                arcazeFromGrid = dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value.ToString();

                            if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                            {
                                isReverse = Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["ledsReverse"].Value);
                                module = Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["Modul"].Value);

                                comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["portLEDs"];
                                port = comboCell.Items.IndexOf(comboCell.Value);

                                comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["pinLEDs"];
                                pin = comboCell.Items.IndexOf(comboCell.Value);

                                resolution = int.Parse(dataGridViewLEDs.Rows[n].Cells["ledResolution"].Value.ToString(), NumberStyles.HexNumber);
                                dataGridViewLEDs.Rows[n].Cells["checkBoxInit"].Value = true;

                                dimmingValue = double.Parse(dataGridViewLEDs.Rows[n].Cells["dimmValue"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);

                                arcazeDeviceIndex = Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["deviceIndex"].Value);

                                if (textBoxIntervalTimer.Text == "")
                                    textBoxIntervalTimer.Text = "20";

                                dimmingQuotient = double.Parse(dataGridViewLEDs.Rows[n].Cells["dimmTime"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);

                                dimmingInterval = double.Parse(textBoxIntervalTimer.Text.Replace(",", "."), CultureInfo.InvariantCulture) /
                                                 (dimmingQuotient * 250);



                                // dimming up
                                if (Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["dimmMode"].Value) == 2)
                                {
                                    dimmingValue += dimmingInterval;

                                    if (dimmingValue >= 1)
                                    {
                                        dimmingValue = 1;
                                        dataGridViewLEDs.Rows[n].Cells["dimmMode"].Value = (1).ToString();
                                    }
                                }
                                else // dimming down
                                {
                                    dimmingValue -= dimmingInterval;

                                    if (dimmingValue <= 0 || dimmingValue > 1)
                                    {
                                        dimmingValue = 0;
                                        dataGridViewLEDs.Rows[n].Cells["dimmMode"].Value = (2).ToString();
                                    }
                                }
                                dataGridViewLEDs.Rows[n].Cells["dimmValue"].Value = dimmingValue.ToString();
                                arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(module, port, pin, resolution, dimmingValue, type, isReverse, checkBoxLog.Checked);

                            }
                            else // switch off
                            {
                                if (double.Parse(dataGridViewLEDs.Rows[n].Cells["dimmValue"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture) > 0.0)
                                {
                                    dataGridViewLEDs.Rows[n].Cells["dimmValue"].Value = (0).ToString();
                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(module, port, pin, resolution, 0, type, isReverse, checkBoxLog.Checked);
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            #region new data

            if (lStateEnabled && !stop && receivedData.Length > 0 && arcazeFound) // New Data ?
            {
                lStateEnabled = false;

                if (checkBoxLog.Checked)
                    ImportExport.LogMessage("Processing package: " + receivedData, true);

                if (receivedData.IndexOf("DAC=stop") != -1)
                    timerstate = State.stop;
                else
                {
                    receivedData += newline;
                    receivedDataBackup = receivedData;
                    GrabValues();

                    if (receivedDataBackup != receivedData)
                    {
                        if (checkBoxLog.Checked)
                            ImportExport.LogMessage("Processing package: " + receivedData, true);

                        receivedData += newline;
                        GrabValues();
                    }
                    dataSetDisplaysLEDs.AcceptChanges();
                }
                receivedData = "";

                lStateEnabled = true;
            }
            #endregion

            #region new joystick button

            if (Joystick.joystickActiv && !stop)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    if (dataSetConfig.Tables["Joysticks"].Rows.Count == 0)
                    {
                        for (int n = 0; n < Joystick.joystickName.Count; n++)
                        {
                            row = dataSetConfig.Tables["Joysticks"].NewRow();
                            row[0] = Joystick.joystickName[n];
                            dataSetConfig.Tables["Joysticks"].Rows.Add(row);
                        }
                    }

                    if (dataSetConfig.Tables["SendKeysFromJoystick"].Rows.Count > 0)
                    {
                        for (int n = 0; n < Joystick.joystickName.Count; n++)
                        {
                            Joystick.UpdateJoystick(n, ref buttonsPressed);

                            if (buttonsPressed != "")
                            {
                                if (checkBoxLog.Checked)
                                    ImportExport.LogMessage("Joystick: " + Joystick.joystickName[n] + " .. processing buttons: --> " + buttonsPressed + " <--", true);

                                rows = new DataRow[] { };
                                rows = dataSetConfig.Tables["SendKeysFromJoystick"].Select("Active=" + true + " AND Joystick='" + Joystick.joystickName[n] + "'");

                                for (int d = 0; d < rows.Length; d++)
                                {
                                    button = rows[d]["Joystickbutton"].ToString();

                                    if (buttonsPressed.IndexOf(button, 0) != -1) // get buttonsPressed and send keystrokes
                                    {
                                        sendKey = rows[d]["Keystrokes"].ToString();
                                        optionKeyCode = rows[d]["OptionKey"].ToString();
                                        processName = rows[d]["Processname"].ToString();

                                        if (sendKey.Trim() != "")
                                        {
                                            if (checkBoxLog.Checked)
                                                ImportExport.LogMessage("Send keystrokes from " + Joystick.joystickName[n] + " OptionKey: " + optionKeyCode + " Key: " + sendKey, true);

                                            if (optionKeyCode.Trim() != "")
                                                InputSimulator.SimulateModifiedKeyStroke(processName, Convert.ToUInt16(optionKeyCode), Convert.ToUInt16(sendKey));
                                            else
                                                InputSimulator.SimulateModifiedKeyStroke(processName, 0, Convert.ToUInt16(sendKey));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    lStateEnabled = true;
                }
            }
            #endregion

            #region Switches - Encoder - Keystrokes

            if (lStateEnabled && !stop)
            {
                lStateEnabled = false;

                #region Switches

                for (int n = 0; n < dataGridViewSwitches.RowCount; n++)
                {
                    if (Convert.ToBoolean(dataGridViewSwitches.Rows[n].Cells["switchesInit"].Value) &&
                        Convert.ToBoolean(dataGridViewSwitches.Rows[n].Cells["switchesActive"].Value))
                    {
                        if (dataGridViewSwitches.Rows[n].Cells["switchesArcaze"].Value != DBNull.Value &&
                            dataGridViewSwitches.Rows[n].Cells["switchesArcaze"].Value != null
                        )
                            arcazeFromGrid = dataGridViewSwitches.Rows[n].Cells["switchesArcaze"].Value.ToString();
                        else
                            arcazeFromGrid = "";

                        if (arcazeFromGrid != "")
                        {
                            if (ActivateArcaze(arcazeFromGrid, false))
                            {
                                numberInputChanged = this.arcazeHid.Command.CmdReadChangedInput();

                                if (numberInputChanged != 0)
                                    textBoxInputChanged.Text = (numberInputChanged < 21 ? "A" + numberInputChanged.ToString() : "B" + (numberInputChanged - 20).ToString());

                                comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesPort"];
                                port = comboCell.Items.IndexOf(comboCell.Value);

                                comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesPin"];
                                pin = comboCell.Items.IndexOf(comboCell.Value);

                                comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesType"];
                                type = comboCell.Items.IndexOf(comboCell.Value);

                                comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesChannel"];
                                channel = comboCell.Items.IndexOf(comboCell.Value);

                                if (dataGridViewSwitches.Rows[n].Cells["switchesReverse"].Value != DBNull.Value)
                                    reverse = Convert.ToBoolean(dataGridViewSwitches.Rows[n].Cells["switchesReverse"].Value);
                                else
                                    reverse = false;

                                if (dataGridViewSwitches.Rows[n].Cells["switchesSendNotOff"].Value != DBNull.Value)
                                    switchesSendNotZero = Convert.ToBoolean(dataGridViewSwitches.Rows[n].Cells["switchesSendNotOff"].Value);
                                else
                                    switchesSendNotZero = false;

                                dezValue = dataGridViewSwitches.Rows[n].Cells["switchesValue"].Value.ToString();

                                if (dataGridViewSwitches.Rows[n].Cells["switchesClickableID"].Value.ToString() != "")
                                {
                                    clickable_ID = Convert.ToInt32(dataGridViewSwitches.Rows[n].Cells["switchesClickableID"].Value);

                                    valueON = dataGridViewSwitches.Rows[n].Cells["switchesValueOn"].Value.ToString().Replace(",", ".");
                                    valueOff = dataGridViewSwitches.Rows[n].Cells["switchesValueOff"].Value.ToString().Replace(",", ".");

                                    rows = new DataRow[] { };
                                    rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("ID=" + clickable_ID);

                                    if (rows.Length > 0)
                                    {
                                        deviceID = Convert.ToInt32(rows[0]["DeviceID"]);
                                        buttonID = Convert.ToInt32(rows[0]["ButtonID"]);

                                        switch (type)
                                        {
                                            case 0: // Switch
                                                ReadPinValue(port, pin, ref pinVal, ref reverse); // On / Off - pinVal = 1 / 0

                                                if (Convert.ToInt32(dezValue) != pinVal)
                                                {
                                                    dataGridViewSwitches.Rows[n].Cells["switchesValue"].Value = pinVal.ToString();

                                                    if (pinVal == 0 && switchesSendNotZero)
                                                        break;

                                                    package = "C" + deviceID + "," + (3000 + buttonID).ToString() + "," + (pinVal == 1 ? valueON : valueOff);

                                                    try
                                                    {
                                                        if (checkBoxLog.Checked)
                                                            ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package, true);

                                                        UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), package);
                                                    }
                                                    catch (Exception f)
                                                    {
                                                        ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package + " ... " + f.ToString(), true);
                                                    }
                                                }
                                                break;

                                            case 1: // Poti
                                                pinVal = ReadADC(channel);

                                                if (Convert.ToInt32(dezValue) != pinVal)
                                                {
                                                    dataGridViewSwitches.Rows[n].Cells["switchesValue"].Value = pinVal.ToString();
                                                    package = "C" + deviceID + "," + (3000 + buttonID).ToString() + "," + pinVal;

                                                    try
                                                    {
                                                        if (checkBoxLog.Checked)
                                                            ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to  IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package, true);

                                                        UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), package);
                                                    }
                                                    catch (Exception f)
                                                    {
                                                        ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package + " ... " + f.ToString(), true);
                                                    }
                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Encoder

                for (int n = 0; n < dataGridViewEncoderValues.RowCount; n++)
                {
                    if (Convert.ToBoolean(dataGridViewEncoderValues.Rows[n].Cells["ActiveEncoder"].Value) &&
                        Convert.ToBoolean(dataGridViewEncoderValues.Rows[n].Cells["InitEncoder"].Value))
                    {
                        if (dataGridViewEncoderValues.Rows[n].Cells["Arcaze"].Value != DBNull.Value &&
                            dataGridViewEncoderValues.Rows[n].Cells["Arcaze"].Value != null
                        )
                            arcazeFromGrid = dataGridViewEncoderValues.Rows[n].Cells["Arcaze"].Value.ToString();
                        else
                            arcazeFromGrid = "";

                        if (arcazeFromGrid != "")
                        {
                            if (ActivateArcaze(arcazeFromGrid, false))
                            {
                                comboCell = (DataGridViewComboBoxCell)dataGridViewEncoderValues.Rows[n].Cells["EncoderNumber"];
                                encoderNo = comboCell.Items.IndexOf(comboCell.Value);

                                if (encoderNo > -1)
                                {
                                    encoderName = dataGridViewEncoderValues.Rows[n].Cells["EncoderNumber"].Value.ToString();

                                    if (dataGridViewEncoderValues.Rows[n].Cells["OldValueEncoder"].Value != DBNull.Value)
                                        encoderOldValue = Convert.ToInt16(dataGridViewEncoderValues.Rows[n].Cells["OldValueEncoder"].Value);
                                    else
                                        encoderOldValue = 3;

                                    encoderDeltaValue = ReadEncoder(encoderNo, ref encoderOldValue);
                                    dataGridViewEncoderValues.Rows[n].Cells["OldValueEncoder"].Value = encoderOldValue;

                                    if (encoderDeltaValue != 0)
                                    {
                                        if (dataGridViewEncoderValues.Rows[n].Cells["ReadValue"].Value != DBNull.Value)
                                            encoderLastValue = Convert.ToInt16(dataGridViewEncoderValues.Rows[n].Cells["ReadValue"].Value);
                                        else
                                            encoderLastValue = 0;

                                        dataGridViewEncoderValues.Rows[n].Cells["ReadValue"].Value = encoderLastValue + encoderDeltaValue;

                                        if (dataGridViewEncoderValues.Rows[n].Cells["DentValue"].Value != DBNull.Value)
                                            encoderDentValue = double.Parse(dataGridViewEncoderValues.Rows[n].Cells["DentValue"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                        else
                                            encoderDentValue = 1;

                                        if (dataGridViewEncoderValues.Rows[n].Cells["MaxValue"].Value != DBNull.Value)
                                            encoderMaxValue = double.Parse(dataGridViewEncoderValues.Rows[n].Cells["MaxValue"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                        else
                                            encoderMaxValue = 10;

                                        if (dataGridViewEncoderValues.Rows[n].Cells["MinValue"].Value != DBNull.Value)
                                            encoderMinValue = double.Parse(dataGridViewEncoderValues.Rows[n].Cells["MinValue"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                        else
                                            encoderMinValue = 0;

                                        encoderCalcValue = encoderDentValue * (encoderLastValue + encoderDeltaValue);
                                        encoderCalcValue = Math.Round(encoderCalcValue, 2);

                                        while (encoderCalcValue > encoderMaxValue) // Limits
                                        {
                                            encoderCalcValue -= (encoderMaxValue - encoderMinValue);
                                        }

                                        while (encoderCalcValue < encoderMinValue) // Limits
                                        {
                                            encoderCalcValue += (encoderMaxValue - encoderMinValue);
                                        }
                                        encoderCalcValue = Math.Round(encoderCalcValue, 2);

                                        if (checkBoxLog.Checked)
                                            ImportExport.LogMessage(GetActiveArcazeName() + " - " + encoderName + " , Value: " + encoderCalcValue.ToString().Replace(",", "."), true);

                                        dataGridViewEncoderValues.Rows[n].Cells["CalcValue"].Value = encoderCalcValue.ToString().Replace(",", ".");
                                        encoderIdentity = Convert.ToInt16(dataGridViewEncoderValues.Rows[n].Cells["EncoderID"].Value);

                                        rows = new DataRow[] { };
                                        rows = dataSetDisplaysLEDs.Tables["MultipostionSwitch"].Select("EncoderID=" + encoderIdentity.ToString());

                                        if (rows.Length > 0)
                                        {
                                            if (Convert.ToBoolean(rows[0]["EveryValue"]))
                                            {
                                                rows[0]["ValueCalc"] = encoderCalcValue;
                                                rows[0]["ValueSend"] = encoderCalcValue;

                                                if (rows[0]["SwitchID"] != DBNull.Value)
                                                {
                                                    MakeDataPackageAndSend(Convert.ToInt32(rows[0]["SwitchID"]), encoderCalcValue);
                                                    dataSetDisplaysLEDs.Tables["MultipostionSwitch"].AcceptChanges();
                                                }
                                            }
                                            else
                                            {
                                                if (rows[0]["ValueCalc"] != DBNull.Value && rows[0]["ValueSend"] != DBNull.Value && rows[0]["SwitchID"] != DBNull.Value)
                                                {
                                                    if (Convert.ToInt32(rows[0]["ValueCalc"]) == encoderCalcValue)
                                                        MakeDataPackageAndSend(Convert.ToInt32(rows[0]["SwitchID"]), double.Parse(rows[0]["ValueSend"].ToString().Replace(",", "."), CultureInfo.InvariantCulture));
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Keystrokes

                for (int n = 0; n < dataGridViewKeystrokes.RowCount; n++)
                {
                    if (Convert.ToBoolean(dataGridViewKeystrokes.Rows[n].Cells["keystrokesActive"].Value) &&
                        Convert.ToBoolean(dataGridViewKeystrokes.Rows[n].Cells["keystrokesInit"].Value))
                    {
                        if (dataGridViewKeystrokes.Rows[n].Cells["keystrokesArcaze"].Value != DBNull.Value &&
                            dataGridViewKeystrokes.Rows[n].Cells["keystrokesArcaze"].Value != null
                        )
                            arcazeFromGrid = dataGridViewKeystrokes.Rows[n].Cells["keystrokesArcaze"].Value.ToString();
                        else
                            arcazeFromGrid = "";

                        if (arcazeFromGrid != "")
                        {
                            if (ActivateArcaze(arcazeFromGrid, false))
                            {
                                comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["keystrokesPort"];
                                port = comboCell.Items.IndexOf(comboCell.Value);

                                comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["keystrokesPin"];
                                pin = comboCell.Items.IndexOf(comboCell.Value);

                                if (dataGridViewKeystrokes.Rows[n].Cells["ValueRead"].Value != DBNull.Value)
                                    dezValue = dataGridViewKeystrokes.Rows[n].Cells["ValueRead"].Value.ToString();
                                else
                                    dezValue = "0";

                                if (dataGridViewKeystrokes.Rows[n].Cells["keystrokesReverse"].Value != DBNull.Value)
                                    reverse = Convert.ToBoolean(dataGridViewKeystrokes.Rows[n].Cells["keystrokesReverse"].Value);
                                else
                                    reverse = false;

                                if (dataGridViewKeystrokes.Rows[n].Cells["SendNotOff"].Value != DBNull.Value)
                                    switchesSendNotZero = Convert.ToBoolean(dataGridViewKeystrokes.Rows[n].Cells["SendNotOff"].Value);
                                else
                                    switchesSendNotZero = false;

                                ReadPinValue(port, pin, ref pinVal, ref reverse);

                                dataGridViewKeystrokes.Rows[n].Cells["ValueRead"].Value = pinVal.ToString();

                                if (pinVal == 0 && switchesSendNotZero)
                                    continue;

                                if (Convert.ToInt32(dezValue) != pinVal)
                                {
                                    comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["Keystrokes"];
                                    sendKey = comboCell.Items.IndexOf(comboCell.Value).ToString();

                                    comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["OptionKeyOne"];
                                    optionKeyCode = comboCell.Items.IndexOf(comboCell.Value).ToString();

                                    comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["OptionKeyTwo"];
                                    optionKeyCodeTwo = comboCell.Items.IndexOf(comboCell.Value).ToString();

                                    processName = dataGridViewKeystrokes.Rows[n].Cells["ProzessName"].Value.ToString();

                                    if (optionKeyCode.Trim() != "" && optionKeyCodeTwo.Trim() != "")
                                    {
                                        if (checkBoxLog.Checked)
                                            ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  OptionKeyOne: " + optionKeyCode + " OptionKeyTwo: " + optionKeyCodeTwo + " Key: " + sendKey, true);

                                        InputSimulator.SimulateModifiedKeyStroke(processName, Convert.ToUInt16(optionKeyCode), Convert.ToUInt16(optionKeyCodeTwo), Convert.ToUInt16(sendKey));
                                    }
                                    else
                                    {
                                        if (optionKeyCode.Trim() != "" || optionKeyCodeTwo.Trim() != "")
                                        {
                                            if (optionKeyCode.Trim() != "")
                                            {
                                                if (checkBoxLog.Checked)
                                                    ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  OptionKeyOne: " + optionKeyCode + " Key: " + sendKey, true);

                                                InputSimulator.SimulateModifiedKeyStroke(processName, Convert.ToUInt16(optionKeyCode), Convert.ToUInt16(sendKey));
                                            }
                                            else
                                            {
                                                if (checkBoxLog.Checked)
                                                    ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  OptionKeyTwo: " + optionKeyCodeTwo + " Key: " + sendKey, true);

                                                InputSimulator.SimulateModifiedKeyStroke(processName, Convert.ToUInt16(optionKeyCodeTwo), Convert.ToUInt16(sendKey));
                                            }
                                        }
                                        else
                                        {
                                            if (checkBoxLog.Checked)
                                                ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  Key: " + sendKey, true);

                                            InputSimulator.SimulateModifiedKeyStroke(processName, 0, Convert.ToUInt16(sendKey));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                lStateEnabled = true;
            }
            #endregion

            GetLogs();

            #region Clean RAM

            loopCounter++;

            if (loopCounter == loopMax)
            {
                MemoryManagement.Reduce();
                loopCounter = 0;
            }
            #endregion
        }

        #region member functions

        public void ShoutHello()
        {
            InputSimulator.SetForegroundProzess("notepad++");
            // Simulate each key stroke
            InputSimulator.SimulateKeyDown(VirtualKeyCode.SHIFT);
            InputSimulator.SimulateKeyPress(VirtualKeyCode.VK_H);
            InputSimulator.SimulateKeyPress(VirtualKeyCode.VK_E);
            InputSimulator.SimulateKeyPress(VirtualKeyCode.VK_L);
            InputSimulator.SimulateKeyPress(VirtualKeyCode.VK_L);
            InputSimulator.SimulateKeyPress(VirtualKeyCode.VK_O);
            InputSimulator.SimulateKeyPress(VirtualKeyCode.VK_1);
            InputSimulator.SimulateKeyUp(VirtualKeyCode.SHIFT);

            // Alternatively you can simulate text entry to acheive the same end result
            //InputSimulator.SimulateTextEntry("HELLO!");
        }

        private String CheckLengthOfDisplayValue(ref DataRow row, String dezValue)
        {
            int displayCount = 0;
            String newValue = "";

            Boolean[] checkBox = new Boolean[8];
            checkBox[0] = Convert.ToBoolean(row["DisplayD0"]);
            checkBox[1] = Convert.ToBoolean(row["DisplayD1"]);
            checkBox[2] = Convert.ToBoolean(row["DisplayD2"]);
            checkBox[3] = Convert.ToBoolean(row["DisplayD3"]);
            checkBox[4] = Convert.ToBoolean(row["DisplayD4"]);
            checkBox[5] = Convert.ToBoolean(row["DisplayD5"]);
            checkBox[6] = Convert.ToBoolean(row["DisplayD6"]);
            checkBox[7] = Convert.ToBoolean(row["DisplayD7"]);

            for (int n = 0; n < 8; n++)
            {
                if (checkBox[n])
                    displayCount++;
            }

            int foundValue = 0;

            for (int n = 0; n < dezValue.Length; n++)
            {
                if (dezValue.Substring(n, 1) != ".")
                    foundValue++;

                if (displayCount >= foundValue)
                    newValue += dezValue.Substring(n, 1);
            }
            return newValue;
        }

        private void ConvertDezHelper(ref String dezValue)
        {
            for (int n = 0; n < 8; n++)
            {
                if (checkBox[n])
                    displayCount++;
            }
            int foundValue = 0;

            for (int n = 0; n < dezValue.Length; n++) // Shorten the value, if to long
            {
                if (dezValue.Substring(n, 1) != ".")
                    foundValue++;

                if (displayCount >= foundValue)
                    newValue += dezValue.Substring(n, 1);
            }
            dezValue = newValue;

            for (int n = 0; n < 8; n++) // Shift the value to the left
            {
                if (!checkBox[n])
                    dezValue += " ";
                else
                    break;
            }
            int numberDigit = 0;

            for (int n = 0; n < dezValue.Length; n++) // Get number of digits
            {
                if (dezValue.Substring(n, 1) != ".")
                    numberDigit++;
            }
            segmentIndex = 0;

            for (int n = 0; n < dezValue.Length; n++)
            {
                if (dezValue.Substring(n, 1) == ".")
                    checkBoxDP[numberDigit - segmentIndex] = true;
                else
                    segmentIndex++;
            }
            newValue = "";

            for (int n = 0; n < dezValue.Length; n++) // Erase the dots
            {
                if (dezValue.Substring(n, 1) != ".")
                    newValue += dezValue.Substring(n, 1);
            }
            dezValue = newValue;

            for (int n = 0; n < 8; n++)
            {
                noValue = false;

                try
                {
                    if (n < dezValue.Length && dezValue.Substring((dezValue.Length - 1) - n, 1) != " ")
                        dezValueByte = Convert.ToInt32(dezValue.Substring((dezValue.Length - 1) - n, 1));
                    else
                    {
                        dezValueByte = 0;
                        noValue = true;
                    }
                }
                catch (Exception f)
                {
                    ImportExport.LogMessage("ConvertDezToSevenSegment ... " + f.ToString(), true);
                    dezValueByte = 0;
                    noValue = true;
                }
                digit[n] = ((checkBox[n] && !noValue) || (checkBox[n] && useLeftPadding && noValue)) ? SevenSegment.GetValuePattern(dezValueByte, checkBoxDP[n]) : SevenSegment.GetNonePattern();
            }
        }

        private void ConvertDezToSevenSegment(ref DataRow row, ref string[] digit, String dezValue)
        {
            if (!Convert.ToBoolean(row["Active"]))
                return;

            //posDot = -1;
            //numberDigit = 0;
            displayCount = 0;
            newValue = "";

            checkBoxDP = new Boolean[8] { false, false, false, false, false, false, false, false };
            checkBox = new Boolean[8];

            checkBox[0] = Convert.ToBoolean(row["DisplayD0"]);
            checkBox[1] = Convert.ToBoolean(row["DisplayD1"]);
            checkBox[2] = Convert.ToBoolean(row["DisplayD2"]);
            checkBox[3] = Convert.ToBoolean(row["DisplayD3"]);
            checkBox[4] = Convert.ToBoolean(row["DisplayD4"]);
            checkBox[5] = Convert.ToBoolean(row["DisplayD5"]);
            checkBox[6] = Convert.ToBoolean(row["DisplayD6"]);
            checkBox[7] = Convert.ToBoolean(row["DisplayD7"]);

            if (dezValue.IndexOf("-", 0) > -1)
            {
                dezValue = dezValue.Substring(dezValue.IndexOf("-", 0) + 1);

                for (int n = 0; n < 8; n++)
                    digit[n] = SevenSegment.GetNonePattern();

                segmentIndex = -1;
            }
            else
            {
                dezValue = dezValue.Replace(",", ".");
                dezValue = CheckLengthOfDisplayValue(ref row, dezValue);

                useLeftPadding = Convert.ToBoolean(row["LeftPadding"]);
                ConvertDezHelper(ref dezValue);
            }
            devAddress = row["ModulID"].ToString();
            maskHex = "";

            for (int n = 7; n > -1; n--)
            {
                maskHex += checkBox[n] ? "1" : "0";
            }

            digitMask = Convert.ToInt32(maskHex, 2).ToString("X");

            if (row["deviceIndex"] != DBNull.Value)
            {
                arcazeDeviceIndex = Convert.ToInt32(row["deviceIndex"]);
                arcazeDevice[arcazeDeviceIndex].WriteDigitsToDisplayDriver(int.Parse(devAddress, NumberStyles.HexNumber), ref digit, int.Parse(digitMask, NumberStyles.HexNumber),
                    checkBoxLog.Checked, Convert.ToBoolean(row["Reverse"]), segmentIndex, Convert.ToInt32(cbRefreshCycle.Text), Convert.ToInt32(cbDelayRefresh.Text));
            }
        }

        private void ConvertDezToSevenSegment(ref string[] digit)
        {
            dezValueByte = 0;
            dezValue = textBoxDezValue.Text.Trim();
            noValue = false;
            //posDot = 0;
            displayCount = 0;
            newValue = "";

            checkBoxDP = new bool[8] { false, false, false, false, false, false, false, false };
            checkBox = new Boolean[8];

            checkBox[0] = checkBoxD0.Checked;
            checkBox[1] = checkBoxD1.Checked;
            checkBox[2] = checkBoxD2.Checked;
            checkBox[3] = checkBoxD3.Checked;
            checkBox[4] = checkBoxD4.Checked;
            checkBox[5] = checkBoxD5.Checked;
            checkBox[6] = checkBoxD6.Checked;
            checkBox[7] = checkBoxD7.Checked;

            if (dezValue.IndexOf("-", 0) > -1)
            {
                dezValue = dezValue.Substring(dezValue.IndexOf("-", 0) + 1);

                for (int n = 0; n < 8; n++)
                    digit[n] = SevenSegment.GetNonePattern();

                segmentIndex = -1;
            }
            else
            {
                dezValue = dezValue.Replace(",", ".");
                ConvertDezHelper(ref dezValue);
            }
            ShowDigits();
            digitMask = GenerateMask();
            textBoxDigitMask.Text = digitMask;
        }

        private void CopyListBoxToClipboard(ListBox lb)
        {
            StringBuilder buffer = new StringBuilder();

            for (int i = 0; i < lb.Items.Count; i++)
            {
                buffer.Append(lb.Items[i].ToString());
                buffer.Append(newline);
            }
            try
            {
                Clipboard.SetText(buffer.ToString());
            }
            catch
            {
                ImportExport.LogMessage("\t > Beim Kopieren ist ins Clipboard ein Fehler aufgetreten. ", false);
            }
        }

        private int DropDownWidth(ComboBox comboBox)
        {
            int maxWidth = 0, temp = 0;
            foreach (var obj in comboBox.Items)
            {
                temp = TextRenderer.MeasureText(obj.ToString(), comboBox.Font).Width;

                if (temp > maxWidth)
                    maxWidth = temp;
            }
            return maxWidth;
        }

        private int DropDownWidth(DataGridViewComboBoxColumn comboBox)
        {
            int maxWidth = 0, temp = 0;
            foreach (var obj in comboBox.Items)
            {
                temp = comboBox.Items.ToString().Length;

                if (temp > maxWidth)
                    maxWidth = temp;
            }
            return maxWidth;
        }

        private string GenerateMask()
        {
            maskHex = "";

            maskHex += checkBoxD7.Checked ? "1" : "0";
            maskHex += checkBoxD6.Checked ? "1" : "0";
            maskHex += checkBoxD5.Checked ? "1" : "0";
            maskHex += checkBoxD4.Checked ? "1" : "0";
            maskHex += checkBoxD3.Checked ? "1" : "0";
            maskHex += checkBoxD2.Checked ? "1" : "0";
            maskHex += checkBoxD1.Checked ? "1" : "0";
            maskHex += checkBoxD0.Checked ? "1" : "0";

            return Convert.ToInt32(maskHex, 2).ToString("X"); ;
        }

        private void GetLogs()
        {
            if (ImportExport.log.Count > 0)
            {
                for (int i = 0; i < ImportExport.log.Count; i++)
                {
                    listBox1.Items.Add(ImportExport.log[i]);

                    if (checkBoxWriteLogsToHD.Checked)
                        listBox3.Items.Add(ImportExport.log[i]);
                }
                try
                {
                    ImportExport.log.Clear();

                    if (checkBoxWriteLogsToHD.Checked)
                        ImportExport.WriteListBoxToFile(listBox3, logFile);

                    listBox3.Items.Clear();

                    listBox1.SelectedIndex = listBox1.Items.Count - 1;
                }
                catch (Exception e)
                {
                    ImportExport.LogMessage("GetLogs ... " + e.ToString(), true);
                }
            }
        }

        private void GetPortPinOfEncoder(ref int encoderNo, ref int port, ref int pin)
        {
            if (encoderNo < 10)
                port = 0;
            else
            {
                port = 1;
                encoderNo -= 10;
            }
            pin = (encoderNo * 2);  // encoder 0 => Pin 0/1; encoder 1 => Pin 2/3; encoder 2 => Pin 4/5 ....
        }

        private string GrabValue(String ID, ref String gotData)
        {
            fragment = "";
            posStart = 0;
            posEnd = 0;
            loadFile = false;

            posStart = gotData.LastIndexOf(":" + ID + "=");

            if (posStart == -1)
            {
                posStart = gotData.LastIndexOf("*" + ID + "=");

                if (posStart == -1)
                    return "";
            }

            //posStart = gotData.IndexOf(ID + "=", 0);
            posStart += 1;

            //if (posStart > -1) // ID found ?
            //{
            fragment = gotData.Substring(posStart); // 44=1.0:ACC01=0.0: ....
            posEnd = fragment.IndexOf(":", 0);

            if (posEnd == -1)
                posEnd = fragment.IndexOf(newline, 0);

            if (posEnd > -1)
            {
                fragment = fragment.Substring(0, posEnd); // 44=1.0
                fragmentID = fragment.Substring(0, fragment.IndexOf("=", 0) - 1);

                if (fragment.IndexOf(searchStringForFile) > -1)
                    loadFile = true;

                if (checkBoxLog.Checked)
                    ImportExport.LogMessage("Processing value: --> " + fragment + " <--", true);

                posStart = fragment.IndexOf("=", 0);
                fragment = fragment.Substring(posStart + 1); // 1.0

                posEnd = fragment.IndexOf(";", 0);

                if (posEnd > -1)
                    fragment = fragment.Substring(0, posEnd);

                fragment = fragment.Trim();

                if (checkBoxLog.Checked)
                    ImportExport.LogMessage("Send       value: --> " + fragment.Trim() + " <--", true);

                if (loadFile)
                {
                    readFile = fragment;
                    timerMain.Interval = timerInterval / 5;
                    timerstate = State.readfile;
                    try
                    {
                        receivedDataBackup = receivedData.Substring(receivedData.IndexOf(searchStringForFile) + 5);
                        receivedData = receivedDataBackup;
                    }
                    catch 
                    {
                        receivedDataBackup = "";
                    }
                    stop = true;
                    return "";
                }
                else
                {
                    digitString = "";

                    for (int n = 0; n < fragment.Length; n++)
                    {
                        characterValue = (byte)Convert.ToChar(fragment.Substring(n, 1));

                        if (characterValue >= 44 && characterValue <= 57 && characterValue != 47)
                            digitString += fragment.Substring(n, 1);
                    }
                    return digitString;
                }
            }
            if (checkBoxLog.Checked)
                ImportExport.LogMessage(" ID --> " + ID + " <-- Value not found.", true);

            return "";
            //}
            //return "";
        }

        private void GrabValues()
        {
            gotData = receivedData;
            arcazeFromGrid = "";

            newValue = GrabValue(searchStringForFile, ref gotData);

            if (loadFile)
                return;

            if (dataGridViewArcaze.Rows.Count > 0)
            {
                #region Displays

                for (int n = 0; n < dataGridViewDisplays.RowCount; n++) // ****** Displays *******
                {
                    if (Convert.ToBoolean(dataGridViewDisplays.Rows[n].Cells["activeDisplaysCheckBox"].Value) &&
                        Convert.ToBoolean(dataGridViewDisplays.Rows[n].Cells["displayInit"].Value))
                    {
                        if (dataGridViewDisplays.Rows[n].Cells["arcazeDisplays"].Value != DBNull.Value &&
                            dataGridViewDisplays.Rows[n].Cells["arcazeDisplays"].Value != null
                        )
                            arcazeFromGrid = dataGridViewDisplays.Rows[n].Cells["arcazeDisplays"].Value.ToString();
                        else
                            arcazeFromGrid = "";

                        if (arcazeFromGrid != "")
                        {
                            if (dataGridViewDisplays.Rows[n].Cells["dcsExportIDDisplays"].Value != DBNull.Value)
                                dcsExportID = dataGridViewDisplays.Rows[n].Cells["dcsExportIDDisplays"].Value.ToString();
                            else
                                dcsExportID = "";

                            if (dcsExportID != "")
                                newValue = GrabValue(dcsExportID, ref gotData); // grab the value
                            else
                                newValue = "";

                            if (newValue != "")
                            {
                                if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                                {
                                    dataGridViewDisplays.Rows[n].Cells["valueDisplays"].Value = newValue;

                                    rows = new DataRow[] { };
                                    rows = dataSetDisplaysLEDs.Tables["Displays"].Select("ID=" + Convert.ToInt32(dataGridViewDisplays.Rows[n].Cells["IDDisplays"].Value).ToString());

                                    ConvertDezToSevenSegment(ref rows[0], ref digit, dataGridViewDisplays.Rows[n].Cells["valueDisplays"].Value.ToString());
                                }
                            }
                        }
                    }
                }
                #endregion

                #region LEDs

                for (int n = 0; n < dataGridViewLEDs.RowCount; n++)
                {
                    if (Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["activeLEDs"].Value) &&
                        Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["checkBoxInit"].Value))
                    {
                        if (dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value != DBNull.Value &&
                            dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value != null
                        )
                            arcazeFromGrid = dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value.ToString();
                        else
                            arcazeFromGrid = "";

                        if (arcazeFromGrid != "")
                        {
                            dcsExportID = dataGridViewLEDs.Rows[n].Cells["dcsExportIDLEDs"].Value.ToString();

                            if (dcsExportID != "")
                                newValue = GrabValue(dcsExportID, ref gotData);
                            else
                                newValue = "";

                            if (newValue != "")
                            {
                                if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                                {
                                    dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value = newValue;

                                    module = Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["Modul"].Value);

                                    if (dataGridViewLEDs.Rows[n].Cells["ledsReverse"].Value == DBNull.Value)
                                        isReverse = false;
                                    else
                                        isReverse = Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["ledsReverse"].Value.ToString());

                                    comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["portLEDs"];
                                    port = comboCell.Items.IndexOf(comboCell.Value);

                                    comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["pinLEDs"];
                                    pin = comboCell.Items.IndexOf(comboCell.Value);

                                    comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["modulTypeID"];

                                    if (comboCell.Value == null)
                                        type = 4;
                                    else
                                        type = comboCell.Items.IndexOf(comboCell.Value) + 1;

                                    if (dataGridViewLEDs.Rows[n].Cells["ledResolution"].Value != DBNull.Value)
                                        resolution = int.Parse(dataGridViewLEDs.Rows[n].Cells["ledResolution"].Value.ToString(), NumberStyles.HexNumber);
                                    else
                                        resolution = 1;

                                    if (dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value != DBNull.Value)
                                    {
                                        pinString = dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value.ToString().Replace(",", ".");

                                        arcazeDeviceIndex = Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["deviceIndex"].Value);

                                        try
                                        {
                                            pinValue = double.Parse(pinString, System.Globalization.CultureInfo.InvariantCulture);

                                            switch (type)
                                            {
                                                case 2:
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(module, port, pin, resolution, Convert.ToDouble(pinValue > 0.5 ? 1 : 0), type, isReverse, checkBoxLog.Checked);
                                                    break;

                                                case 3:
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(module, port, pin, resolution, double.Parse(pinString, System.Globalization.CultureInfo.InvariantCulture), type, isReverse, checkBoxLog.Checked);
                                                    break;

                                                case 4:
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(module, port, pin, resolution, Convert.ToDouble(pinValue > 0.5 ? 1 : 0), type, isReverse, checkBoxLog.Checked);
                                                    break;
                                            }
                                        }
                                        catch
                                        {
                                            ImportExport.LogMessage("LED Value: " + dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value + " .... out of range ...", true);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
            }
        }

        private void InitDrivers()
        {
            if (arcazeDevice != null)
                arcazeDevice.Clear();

            InitDisplaysDrivers();
            InitLEDDrivers();

            SetPinDirectionForLEDs();
            InitSwitches();
            InitEncoders();
            InitKeystrokes();
        }

        private void InitAll()
        {
            FindAllArcaze();
            dataSetConfig.Tables["Joysticks"].Clear();

            if (!arcazeFound)
                timerstate = State.startup;
            else
                timerstate = State.init;
        }

        private void InitConfig()
        {
            if (dataSetConfig.Tables["Config"].Rows.Count > 0)
            {
                textBoxLastFile.Text = dataSetConfig.Tables["Config"].Rows[0]["LastFile"].ToString();
                textBoxIP.Text = dataSetConfig.Tables["Config"].Rows[0]["IP"].ToString();
                textBoxPortListener.Text = dataSetConfig.Tables["Config"].Rows[0]["PortListener"].ToString();

                brightness = int.Parse(dataSetConfig.Tables["Config"].Rows[0]["Tlc5943Brightness"].ToString(), NumberStyles.HexNumber);
                trackBarLEDDriverBrightness.Value = (brightness > 127 ? 127 : brightness);

                brightness = int.Parse(dataSetConfig.Tables["Config"].Rows[0]["Display7219DigitsIntensity"].ToString(), NumberStyles.HexNumber);
                trackBarDisplayBrightness.Value = (brightness > 15 ? 15 : brightness);

                cbDelayRefresh.Text = dataSetConfig.Tables["Config"].Rows[0]["refreshDelay"].ToString();

                if (cbDelayRefresh.Text == "")
                    cbDelayRefresh.Text = "10";

                cbRefreshCycle.Text = dataSetConfig.Tables["Config"].Rows[0]["refreshCycles"].ToString();

                if (cbRefreshCycle.Text == "")
                    cbRefreshCycle.Text = "2";

                ImportExport.LogMessage("Used configuration file '" + textBoxLastFile.Text + "'", true);

                comboBoxLEDValue.SelectedValue = 1;
                this.Refresh();
            }
        }

        private void InitDisplaysDrivers()
        {
            DataRow[] rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["Displays"].Select("Active=" + true, "Arcaze ASC");

            for (int n = 0; n < rows.Length; n++)
            {
                rows[n]["Init"] = false;

                if (rows[n]["Arcaze"] != DBNull.Value)
                    arcazeFromGrid = rows[n]["Arcaze"].ToString();
                else
                    arcazeFromGrid = "";

                if (arcazeFromGrid != "")
                {
                    if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                    {
                        arcazeDeviceFound = false;

                        for (int m = 0; m < arcazeDevice.Count; m++)
                        {
                            if (arcazeDevice[m].GetSerial == arcazeHid.Info.Serial)
                            {
                                arcazeDeviceFound = true;
                                arcazeDeviceIndex = m;
                                rows[n]["deviceIndex"] = arcazeDeviceIndex;

                                arcazeDevice[arcazeDeviceIndex].InitDisplayDriver(Convert.ToInt32(rows[n]["ModulID"]), 0, trackBarDisplayBrightness.Value, 8);

                                rows[n]["Init"] = true;
                            }
                        }

                        if (!arcazeDeviceFound)
                        {
                            arcazeDevice.Add(new ArcazeDevice(arcazeHid.Info));
                            arcazeDeviceIndex = arcazeDevice.Count - 1;
                            rows[n]["deviceIndex"] = arcazeDeviceIndex;

                            arcazeDevice[arcazeDeviceIndex].InitDisplayDriver(Convert.ToInt32(rows[n]["ModulID"]), 0, trackBarDisplayBrightness.Value, 8);

                            rows[n]["Init"] = true;
                        }
                    }
                }
            }
            dataSetDisplaysLEDs.AcceptChanges();
        }

        private void InitTables()
        {
            dataSetConfig.Tables["Joysticks"].Clear();

            dataSetConfig.Tables["VirtualOptionKeys"].Clear();

            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add(" ", "0x00");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("SHIFT", "0x10");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("CTRL", "0x11");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("ALT", "0x12");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("CAPS LOCK", "0x14");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("LWIN", "0x5B");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("RWIN", "0x5C");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("LSHIFT", "0xA0");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("RSHIFT", "0xA1");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("LCTRL", "0xA2");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("RCTRL", "0xA3");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("LALT", "0xA4");
            dataSetConfig.Tables["VirtualOptionKeys"].Rows.Add("RALT", "0xA5");

            dataSetConfig.Tables["VirtualKeys"].Clear();

            dataSetConfig.Tables["VirtualKeys"].Rows.Add("", "0x00");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("BACKSPACE", "0x08");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("TAB", "0x09");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("ENTER", "0x0D");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("PAUSE", "0x13");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("SPACEBAR", "0x20");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("ESC", "0x1B");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("PAGE UP", "0x21");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("PAGE DOWN", "0x22");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("END", "0x23");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("HOME", "0x24");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("LEFT ARROW", "0x25");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("RIGHT ARROW", "0x27");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("UP ARROW", "0x26");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("DOWN ARROW", "0x28");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("INS", "0x2D");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("DEL", "0x2E");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("0", "0x30");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("1", "0x31");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("2", "0x32");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("3", "0x33");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("4", "0x34");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("5", "0x35");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("6", "0x36");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("7", "0x37");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("8", "0x38");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("9", "0x39");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("A", "0x41");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("B", "0x42");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("C", "0x43");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("D", "0x44");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("E", "0x45");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F", "0x46");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("G", "0x47");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("H", "0x48");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("I", "0x49");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("J", "0x4A");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("K", "0x4B");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("L", "0x4C");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("M", "0x4D");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("N", "0x4E");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("O", "0x4F");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("P", "0x50");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("Q", "0x51");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("R", "0x52");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("S", "0x53");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("T", "0x54");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("U", "0x55");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("V", "0x56");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("W", "0x57");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("X", "0x58");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("Y", "0x59");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("Z", "0x5A");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 0", "0x60");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 1", "0x61");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 2", "0x62");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 3", "0x63");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 4", "0x64");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 5", "0x65");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 6", "0x66");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 7", "0x67");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 8", "0x68");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMPAD 9", "0x69");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("MULTIPLY", "0x6A");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("ADD", "0x6B");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("SEPARATOR", "0x6C");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("SUBTRACT", "0x6D");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("DECIMAL", "0x6E");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("DIVIDE", "0x6F");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F1", "0x70");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F2", "0x71");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F3", "0x72");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F4", "0x73");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F5", "0x74");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F6", "0x75");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F7", "0x76");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F8", "0x77");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F9", "0x78");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F10", "0x79");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F11", "0x7A");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("F12", "0x7B");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("NUMLOCK", "0x90");
            dataSetConfig.Tables["VirtualKeys"].Rows.Add("SCROLL", "0x91");

            dataSetConfig.AcceptChanges();


            for (int n = 0; n < dataSetDisplaysLEDs.Tables["Arcaze"].Rows.Count; n++)
            {
                dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["Active"] = false;

                if (dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["ModulTypID"].ToString() == "2") // repair
                    dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["ModulTypID"] = "LED-Driver 2";
                if (dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["ModulTypID"].ToString() == "3")
                    dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["ModulTypID"] = "LED-Driver 3";
                if (dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["ModulTypID"].ToString() == "4")
                    dataSetDisplaysLEDs.Tables["Arcaze"].Rows[n]["ModulTypID"] = "Arcaze USB";
            }

            for (int n = 0; n < dataSetDisplaysLEDs.Tables["Displays"].Rows.Count; n++)
            {
                dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["Init"] = false;

                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "0")  // repair
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "00";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "1")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "01";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "2")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "02";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "3")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "03";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "4")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "04";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "5")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "05";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "6")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "06";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "7")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "07";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "8")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "08";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "9")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "09";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "10")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "0A";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "11")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "0B";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "12")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "0C";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "13")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "0D";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "14")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "0E";
                if (dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"].ToString() == "15")
                    dataSetDisplaysLEDs.Tables["Displays"].Rows[n]["ModulID"] = "0F";
            }

            for (int n = 0; n < dataSetDisplaysLEDs.Tables["LEDs"].Rows.Count; n++)
            {
                dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["Init"] = false;
                dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["dimmValue"] = (0).ToString();

                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulTypeID"].ToString() == "2")  // repair
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulTypeID"] = "LED-Driver 2";
                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulTypeID"].ToString() == "3")
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulTypeID"] = "LED-Driver 3";
                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulTypeID"].ToString() == "4")
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulTypeID"] = "Arcaze USB";

                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"].ToString() == "0")
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"] = "0";
                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"].ToString() == "1")
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"] = "1";
                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"].ToString() == "2")
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"] = "2";
                if (dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"].ToString() == "3")
                    dataSetDisplaysLEDs.Tables["LEDs"].Rows[n]["ModulID"] = "3";
            }

            for (int n = 0; n < dataSetDisplaysLEDs.Tables["Switches"].Rows.Count; n++)
            {
                dataSetDisplaysLEDs.Tables["Switches"].Rows[n]["Init"] = false;
                dataSetDisplaysLEDs.Tables["Switches"].Rows[n]["ValueDez"] = (0).ToString();
            }

            for (int n = 0; n < dataSetDisplaysLEDs.Tables["Encoder"].Rows.Count; n++)
            {
                dataSetDisplaysLEDs.Tables["Encoder"].Rows[n]["Init"] = false;
            }

            for (int n = 0; n < dataSetDisplaysLEDs.Tables["Keystrokes"].Rows.Count; n++)
            {
                dataSetDisplaysLEDs.Tables["Keystrokes"].Rows[n]["Init"] = false;
            }

            dataSetDisplaysLEDs.AcceptChanges();
        }

        private void InitEncoders()
        {
            for (int n = 0; n < dataGridViewEncoderValues.RowCount; n++)
            {
                dataGridViewEncoderValues.Rows[n].Cells["InitEncoder"].Value = false;
                dataGridViewEncoderValues.Rows[n].Cells["ReadValue"].Value = 0;
                dataGridViewEncoderValues.Rows[n].Cells["CalcValue"].Value = 0;

                if (dataGridViewEncoderValues.Rows[n].Cells["Arcaze"].Value != DBNull.Value &&
                    dataGridViewEncoderValues.Rows[n].Cells["Arcaze"].Value != null
                )
                    arcazeFromGrid = dataGridViewEncoderValues.Rows[n].Cells["Arcaze"].Value.ToString();
                else
                    arcazeFromGrid = "";

                if (arcazeFromGrid != "")
                {
                    if (ActivateArcaze(arcazeFromGrid, false))
                    {
                        comboCell = (DataGridViewComboBoxCell)dataGridViewEncoderValues.Rows[n].Cells["EncoderNumber"];
                        encoderNo = comboCell.Items.IndexOf(comboCell.Value);

                        if (encoderNo > -1)
                        {
                            GetPortPinOfEncoder(ref encoderNo, ref port, ref pin);

                            ImportExport.LogMessage("Init Encoder " + (encoderNo + 1).ToString("D2"), true);

                            SetPinDirection(port, pin, 0);          // direction input
                            SetPinDirection(port, pin + 1, 0);

                            dataGridViewEncoderValues.Rows[n].Cells["InitEncoder"].Value = true;
                            dataGridViewEncoderValues.Rows[n].Cells["OldValueEncoder"].Value = 3;
                        }
                    }
                }
            }
        }

        private void InitKeystrokes()
        {
            for (int n = 0; n < dataGridViewKeystrokes.RowCount; n++) // ******** Keystrokes ********
            {
                dataGridViewKeystrokes.Rows[n].Cells["keystrokesInit"].Value = false;

                if (dataGridViewKeystrokes.Rows[n].Cells["keystrokesArcaze"].Value != DBNull.Value &&
                    dataGridViewKeystrokes.Rows[n].Cells["keystrokesArcaze"].Value != null
                )
                    arcazeFromGrid = dataGridViewKeystrokes.Rows[n].Cells["keystrokesArcaze"].Value.ToString();
                else
                    arcazeFromGrid = "";

                if (arcazeFromGrid != "")
                {
                    if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                    {
                        comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["keystrokesPort"];
                        port = comboCell.Items.IndexOf(comboCell.Value);

                        comboCell = (DataGridViewComboBoxCell)dataGridViewKeystrokes.Rows[n].Cells["keystrokesPin"];
                        pin = comboCell.Items.IndexOf(comboCell.Value);

                        SetPinDirection(port, pin, 0); // direction input
                        dataGridViewKeystrokes.Rows[n].Cells["keystrokesInit"].Value = true;
                        reverse = Convert.ToBoolean(dataGridViewSwitches.Rows[n].Cells["switchesReverse"].Value);

                        ReadPinValue(port, pin, ref pinVal, ref reverse);

                        dataGridViewKeystrokes.Rows[n].Cells["ValueRead"].Value = pinVal.ToString();
                    }
                }
            }
        }

        private int GetArcazeInstanceFromDisplays(string arcaze)
        {
            rows = new DataRow[] { };

            rows = dataSetDisplaysLEDs.Tables["Displays"].Select("Init=" + true + " AND Arcaze ='" + arcaze + "'");

            if (rows.Length > 0)
            {
                try
                {
                    return Convert.ToInt32(rows[0]["deviceIndex"]);
                }
                catch
                {
                    return -1;
                }
            }
            else
            {
                if (arcaze != "")
                {
                    if (ActivateArcaze(arcaze, false))
                    {
                        for (int m = 0; m < arcazeDevice.Count; m++)
                        {
                            if (arcazeDevice[m].GetSerial == arcazeHid.Info.Serial)
                                return m;
                        }
                        arcazeDevice.Add(new ArcazeDevice(arcazeHid.Info));
                        return (arcazeDevice.Count - 1);
                    }
                }
                return -1;
            }
        }

        private int GetArcazeInstanceFromLEDs(string arcaze)
        {
            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["LEDs"].Select("Init=" + true + " AND Arcaze ='" + arcaze + "'");

            if (rows.Length > 0)
            {
                try
                {
                    return Convert.ToInt32(rows[0]["deviceIndex"]);
                }
                catch
                {
                    return -1;
                }
            }
            else
            {
                if (arcaze != "")
                {
                    if (ActivateArcaze(arcaze, false))
                    {
                        for (int m = 0; m < arcazeDevice.Count; m++)
                        {
                            if (arcazeDevice[m].GetSerial == arcazeHid.Info.Serial)
                                return m;
                        }
                        arcazeDevice.Add(new ArcazeDevice(arcazeHid.Info));
                        return (arcazeDevice.Count - 1);
                    }
                }
                return -1;
            }
        }

        private bool InitLEDDriver()
        {
            comboBoxModuleType.Refresh();
            arcazeDeviceIndex = GetArcazeInstanceFromLEDs(ComboBoxArcaze.Text);

            if (arcazeDeviceIndex > -1)
            {
                switch (comboBoxModuleType.Text)
                {
                    case "":
                        break;

                    case "Display-Driver":
                        break;

                    case "LED-Driver 2":
                        return arcazeDevice[arcazeDeviceIndex].InitExtensionPort(ArcazeCommand.ExtModuleType.LedDriver2, comboBoxModul.SelectedIndex, 1, trackBarLEDDriverBrightness.Value);

                    case "LED-Driver 3":
                        return arcazeDevice[arcazeDeviceIndex].InitExtensionPort(ArcazeCommand.ExtModuleType.LedDriver3, comboBoxModul.SelectedIndex, 8, trackBarLEDDriverBrightness.Value);

                    case "Arcaze USB":
                        return true;
                }
            }
            return false;
        }

        private void InitLEDDriver(string moduleTyp, int module, int resolution, int brightness)
        {
            switch (moduleTyp)
            {
                case "LED-Driver 2":
                    arcazeDevice[arcazeDeviceIndex].InitExtensionPort(ArcazeCommand.ExtModuleType.LedDriver2, module, resolution, brightness);
                    break;

                case "LED-Driver 3":
                    arcazeDevice[arcazeDeviceIndex].InitExtensionPort(ArcazeCommand.ExtModuleType.LedDriver3, module, resolution, brightness);
                    break;
            }
        }

        private bool InitLEDDrivers()
        {
            DataRow[] rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["LEDs"].Select("Active=" + true, "Arcaze ASC, ModulID ASC");

            for (int n = 0; n < rows.Length; n++)
            {
                rows[n]["Init"] = false;

                if (rows[n]["Arcaze"] != DBNull.Value)
                    arcazeFromGrid = rows[n]["Arcaze"].ToString();
                else
                    arcazeFromGrid = "";

                if (arcazeFromGrid != "")
                {
                    if (ActivateArcaze(arcazeFromGrid, false))
                    {
                        moduleTyp = rows[n]["modulTypeID"].ToString();
                        module = Convert.ToInt32(rows[n]["ModulID"]);
                        resolution = int.Parse(rows[n]["Resolution"].ToString(), NumberStyles.HexNumber);

                        arcazeDeviceFound = false;

                        for (int m = 0; m < arcazeDevice.Count; m++)
                        {
                            if (arcazeDevice[m].GetSerial == arcazeHid.Info.Serial)
                            {
                                arcazeDeviceFound = true;

                                arcazeDeviceIndex = m;
                                rows[n]["deviceIndex"] = arcazeDeviceIndex;

                                InitLEDDriver(moduleTyp, module, resolution, trackBarLEDDriverBrightness.Value);

                                rows[n]["Init"] = true;
                            }
                        }

                        if (!arcazeDeviceFound)
                        {
                            arcazeDevice.Add(new ArcazeDevice(arcazeHid.Info));
                            arcazeDeviceIndex = arcazeDevice.Count - 1;
                            rows[n]["deviceIndex"] = arcazeDeviceIndex;

                            InitLEDDriver(moduleTyp, module, resolution, trackBarLEDDriverBrightness.Value);

                            rows[n]["Init"] = true;
                        }
                    }
                }
            }
            dataSetDisplaysLEDs.AcceptChanges();
            return true;
        }

        private void SetPinDirectionForLEDs()
        {
            for (int n = 0; n < dataGridViewLEDs.RowCount; n++) // ******** LEDs ********
            {
                if (Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["checkBoxInit"].Value))
                {
                    if (dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value != DBNull.Value &&
                        dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value != null
                    )
                        arcazeFromGrid = dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value.ToString();
                    else
                        arcazeFromGrid = "";

                    if (arcazeFromGrid != "")
                    {
                        if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                        {
                            comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["modulTypeID"];

                            if (comboCell.Value == null)
                                type = 4;
                            else
                                type = comboCell.Items.IndexOf(comboCell.Value) + 1;

                            comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["portLEDs"];
                            port = comboCell.Items.IndexOf(comboCell.Value);

                            comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["pinLEDs"];
                            pin = comboCell.Items.IndexOf(comboCell.Value);

                            arcazeDeviceIndex = Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["deviceIndex"].Value);

                            switch (type)
                            {
                                case 0: // Internal
                                    break;

                                case 1: // Display-Driver
                                    break;

                                case 2: // LED-Driver 2
                                    break;

                                case 3: // LED-Driver 3
                                    break;

                                case 4: // Arcaze USB
                                    arcazeDevice[arcazeDeviceIndex].SetPinDirection(port, pin, 1); // direction output
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private void InitSwitches()
        {
            for (int n = 0; n < dataGridViewSwitches.RowCount; n++) // ******** Switches ********
            {
                dataGridViewSwitches.Rows[n].Cells["switchesInit"].Value = false;

                if (dataGridViewSwitches.Rows[n].Cells["switchesArcaze"].Value != DBNull.Value &&
                    dataGridViewSwitches.Rows[n].Cells["switchesArcaze"].Value != null
                )
                    arcazeFromGrid = dataGridViewSwitches.Rows[n].Cells["switchesArcaze"].Value.ToString();
                else
                    arcazeFromGrid = "";

                if (arcazeFromGrid != "")
                {
                    if (ActivateArcaze(arcazeFromGrid, false)) // Activate the arcaze, if necessary
                    {
                        comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesPort"];
                        port = comboCell.Items.IndexOf(comboCell.Value);

                        comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesPin"];
                        pin = comboCell.Items.IndexOf(comboCell.Value);

                        comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesType"];
                        type = comboCell.Items.IndexOf(comboCell.Value);

                        dataGridViewSwitches.Rows[n].Cells["switchesInit"].Value = true;

                        switch (type)
                        {
                            case 0: // Switch
                                SetPinDirection(port, pin, 0); // direction input
                                break;
                            case 1: // Poti
                                break;
                            case 2: // Encoder
                                break;
                        }
                    }
                }
            }
        }

        private void MakeDataPackageAndSend(int switchIdentity, double sendValue)
        {
            clickableRows = new DataRow[] { };
            clickableRows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("ID=" + switchIdentity.ToString());

            if (clickableRows.Length > 0)
            {
                try
                {
                    deviceID = Convert.ToInt32(clickableRows[0]["DeviceID"]);
                    buttonID = Convert.ToInt32(clickableRows[0]["ButtonID"]);

                    package = "C" + deviceID + "," + (3000 + buttonID).ToString() + "," + sendValue.ToString().Replace(",", ".");

                    if (checkBoxLog.Checked)
                        ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to  IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package, true);

                    UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), package);
                }
                catch (Exception f)
                {
                    ImportExport.LogMessage("Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package + " ... " + f.ToString(), true);
                }
            }
        }

        private uint ResolutionValue(int resolution)
        {
            switch (resolution)
            {
                case 2:
                    return 3;
                case 4:
                    return 15;
                case 8:
                    return 255;
                case 10:
                    return 1023;
                default:
                    return 1;
            }
        }

        private void SaveConfig()
        {
            if (dataSetConfig.Tables["Config"].Rows.Count > 0)
            {
                dataSetConfig.Tables["Config"].Rows[0]["LastFile"] = textBoxLastFile.Text;
                dataSetConfig.Tables["Config"].Rows[0]["WriteToHD"] = checkBoxWriteLogsToHD.Checked;
                dataSetConfig.Tables["Config"].Rows[0]["LogAllActions"] = checkBoxLog.Checked;
                dataSetConfig.Tables["Config"].Rows[0]["PortListener"] = textBoxPortListener.Text;
                dataSetConfig.Tables["Config"].Rows[0]["PortSender"] = textBoxPortSender.Text;
                dataSetConfig.Tables["Config"].Rows[0]["IntervalTimer"] = textBoxIntervalTimer.Text;
                dataSetConfig.Tables["Config"].Rows[0]["Tlc5943Brightness"] = trackBarLEDDriverBrightness.Value.ToString("X2");
                dataSetConfig.Tables["Config"].Rows[0]["Display7219DigitsIntensity"] = trackBarDisplayBrightness.Value.ToString("X2");
                dataSetConfig.Tables["Config"].Rows[0]["TestData"] = textBoxTestDataPackage.Text.Replace(newline, "");
                dataSetConfig.Tables["Config"].Rows[0]["refreshDelay"] = cbDelayRefresh.Text;
                dataSetConfig.Tables["Config"].Rows[0]["refreshCycles"] = cbRefreshCycle.Text;
            }
            dataSetConfig.Tables["Config"].AcceptChanges();
            ImportExport.DatasetToXml("config.xml", dataSetConfig);
        }

        private void SendKeystrokesTimeBased(ref DataRow[] rows)
        {
            if (rows.Length > 0)
            {
                for (int n = 0; n < rows.Length; n++)
                {
                    try
                    {
                        sendKey = rows[n]["Keystrokes"].ToString();
                        optionKeyCode = rows[n]["OptionKey"].ToString();
                        processName = rows[n]["Processname"].ToString();

                        sendKeyDelay = double.Parse(rows[n]["Delay"].ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                        sendKeyValue = double.Parse(rows[n]["Value"].ToString().Replace(",", "."), CultureInfo.InvariantCulture);

                        sendKeyValue += double.Parse((timerMain.Interval).ToString().Replace(",", "."), CultureInfo.InvariantCulture) / 1000;

                        if (sendKeyValue >= sendKeyDelay)
                        {
                            rows[n]["Value"] = sendKeyDelay;

                            if (sendKey != "")
                            {
                                ImportExport.LogMessage("Send keystrokes: " + sendKey, true);

                                if (optionKeyCode.Trim() != "")
                                    InputSimulator.SimulateModifiedKeyStroke(processName, Convert.ToUInt16(optionKeyCode), Convert.ToUInt16(sendKey));
                                else
                                    InputSimulator.SimulateModifiedKeyStroke(processName, 0, Convert.ToUInt16(sendKey));
                            }
                            rows[n]["Sended"] = true;
                            dataSetConfig.Tables["SendKeysTimeBased"].AcceptChanges();
                        }
                        rows[n]["Value"] = sendKeyValue;
                    }
                    catch (Exception f)
                    {
                        ImportExport.LogMessage("Send keystrokes: " + sendKey + " ... " + f.ToString(), true);
                    }
                }
            }
            else
                sendOpenKeysTimeBased = false;
        }

        private void ShowDigits()
        {
            textBoxDigit0.Text = digit[0];
            textBoxDigit1.Text = digit[1];
            textBoxDigit2.Text = digit[2];
            textBoxDigit3.Text = digit[3];
            textBoxDigit4.Text = digit[4];
            textBoxDigit5.Text = digit[5];
            textBoxDigit6.Text = digit[6];
            textBoxDigit7.Text = digit[7];
            this.Refresh();
        }

        private void StartListener()
        {
            UDP.StartListener(Convert.ToInt16(textBoxPortListener.Text.Trim()), ref receivedData);
        }

        private void SwitchInverseOff()
        {
            dataSetDisplaysLEDs.AcceptChanges();

            DataRow[] rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["LEDs"].Select("Reverse=" + true, " Arcaze ASC, Pin ASC");

            for (int n = 0; n < rows.Length; n++)
            {
                try
                {
                    dcsExportID = rows[n]["DCSExportID"].ToString();

                    if (dcsExportID != null && dcsExportID != "")
                    {
                        receivedData = ":" + dcsExportID + "= 1:";

                        GrabValues();
                        receivedData = "";
                        rows[n]["Value"] = "0";
                    }
                }
                catch (Exception e)
                {
                    ImportExport.LogMessage("Switch all Off, " + dcsExportID + " ... " + e.ToString(), true);
                }
            }

        }

        private void SwitchAll(bool onValue)
        {
            dataSetDisplaysLEDs.AcceptChanges();
            stop = true;

            DataRow[] rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["LEDs"].Select("Active=" + true, " Arcaze ASC, Pin ASC");

            for (int n = 0; n < rows.Length; n++)
            {
                try
                {
                    dcsExportID = rows[n]["DCSExportID"].ToString();

                    if (dcsExportID != null && dcsExportID != "")
                    {
                        receivedData = ":" + dcsExportID + "=" + (onValue ? "1" : "0") + ":";

                        GrabValues();
                        receivedData = "";
                    }
                }
                catch (Exception e)
                {
                    ImportExport.LogMessage("Switch all " + (onValue ? "On" : "Off") + ", " + dcsExportID + " ... " + e.ToString(), true);
                }
            }

            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["Displays"].Select("Active=" + true, " Arcaze ASC");

            for (int n = 0; n < rows.Length; n++)
            {
                dcsExportID = rows[n]["DCSExportID"].ToString();

                try
                {
                    if (dcsExportID != null && dcsExportID != "")
                    {
                        receivedData = ":" + dcsExportID + "=" + (onValue ? "88888888" : "-") + ":";
                        GrabValues();
                        receivedData = "";
                    }
                }
                catch (Exception e)
                {
                    ImportExport.LogMessage("SwitchAll " + (onValue ? "On" : "Off") + " ... " + e.ToString(), true);
                }
            }
            stop = false;
            buttonStop.Text = (stop ? "Start" : "Stop");
            buttonStop.ForeColor = (stop ? Color.Green : Color.Red);
        }

        private void SystemCheck()
        {
            buttonSendTestPattern.Enabled = false;
            testLoop = 0;
            ledOn = !ledOn;
            labelOnOff.Text = (ledOn ? "Switch 'Off'" : "Switch 'On'");
            timerstate = State.checkLEDS;
        }

        #endregion

        #region Arcaze

        private bool ActivateArcaze(String arcazeNew, bool log = true)
        {
            if (arcazeNew == "")
                return false;

            if (arcazeAddress == arcazeNew)
                return true;

            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["Arcaze"].Select("Arcaze='" + arcazeNew + "'");

            if (rows.Length == 0)
                return false;

            for (int n = 0; n < rows.Length; n++)
            {
                arcazeActive = Convert.ToBoolean(rows[n]["Active"]);

                if (!arcazeActive)
                    return false;
            }

            if (arcazeNew != arcazeAddress)
            {
                for (int n = 0; n < presentArcaze.Count; n++)
                {
                    if (arcazeNew == presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")")
                    {
                        try
                        {
                            if (checkBoxLog.Checked && arcazeHid.Info.Connected && log)
                                ImportExport.LogMessage("Disconnect " + arcazeHid.Info.DeviceName + " (" + arcazeHid.Info.Serial + ")", true);

                            if (arcazeHid.Info.Connected)
                                arcazeHid.Disconnect();

                            if (checkBoxLog.Checked && log)
                                ImportExport.LogMessage("Connect " + presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")", true);

                            this.arcazeHid.Connect(presentArcaze[n].Path);

                            arcazeAddress = arcazeNew;
                        }
                        catch (Exception e)
                        {
                            ImportExport.LogMessage("Cannot switch to Arcaze .. " + presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ") .. " + e.ToString(), true);
                        }
                        return true;
                    }
                }
                return false;
            }
            else
                return true;
        }

        private void ChangeArcase(string presentArcazePath)
        {
            if (arcazeHid.Info.Connected)
                arcazeHid.Disconnect();

            if (checkBoxLog.Checked)
                ImportExport.LogMessage("Connect " + presentArcazePath, true);

            arcazeHid.Connect(presentArcazePath);
        }

        private void DisconnectArcaze()
        {
            if (arcazeHid.Info.Connected)
                arcazeHid.Disconnect();
        }

        private ArcazeCommand.OutputOperators getOutputOperator()
        {
            // Read output operator from combo box
            //switch (comboBoxCCOutputOperator.SelectedItem.ToString())
            //{
            //    case "And":
            //        return ArcazeCommand.OutputOperators.And;
            //    case "Or":
            //        return ArcazeCommand.OutputOperators.Or;
            //    case "Xor":
            //        return ArcazeCommand.OutputOperators.Xor;
            //    case "PlainWrite":
            //    default:
            return ArcazeCommand.OutputOperators.PlainWrite;
            //}
        }

        private bool FindAllArcaze()
        {
            allArcaze = new List<DeviceInfo>(8);

            arcazeHid.Find(allArcaze);
            presentArcaze = arcazeHid.RemoveSameSerialDevices(allArcaze);

            ComboBoxArcaze.Items.Clear();
            ComboBoxArcaze.Text = "";

            //            row = dataSetDisplaysLEDs.Tables["Arcaze"].NewRow();
            findArcaze = new DataRow[] { };
            dataSetDisplaysLEDs.Tables["Arcaze"].Clear();

            ComboBoxArcaze.Text = "";

            for (int n = 0; n < presentArcaze.Count; n++)
            {
                ComboBoxArcaze.Items.Add(presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")");

                row = dataSetDisplaysLEDs.Tables["Arcaze"].NewRow();
                row[0] = presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")";
                row[1] = true;
                try
                {
                    dataSetDisplaysLEDs.Tables["Arcaze"].Rows.Add(row);
                }
                catch { }
            }
            dataSetDisplaysLEDs.Tables["Arcaze"].AcceptChanges();

            if (presentArcaze.Count > 0)
            {
                ImportExport.LogMessage(presentArcaze.Count.ToString() + " Arcaze found .. ", true);

                ComboBoxArcaze.SelectedIndex = 0;
                arcazeHid.Connect(presentArcaze[0].Path);
            }
            return (presentArcaze.Count > 0);
        }

        private string GetActiveArcazeName()
        {
            return arcazeHid.Info.DeviceName + " (" + arcazeHid.Info.Serial + ")";
        }

        private void ResetArcaze()
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                return;

            try
            {
                if (checkBoxLog.Checked)
                    ImportExport.LogMessage("CmdReset() ", true);

                arcazeHid.Command.CmdReset();
            }
            catch (Exception e)
            {
                ImportExport.LogMessage("arcazeHid.ArcazeCommands.Reset() .. " + e.ToString(), true);
            }
        }

        #endregion

        #region Display

        /// <summary>
        /// Init a display driver
        /// </summary>
        /// <param name="devAdress">The unique device address of the Display Module (set by the rotary switch on the board)</param>
        /// <param name="decodeMode">0x00 = No decoding / 0xFF = Code B Decoding (only lower data nibble used then), default = 0x00</param>
        /// <param name="intensity">0x00 ... 0x0F</param>
        /// <param name="scanLimit">4 ... 8 allowed (default = 8, no need to change)</param>
        //private void InitDisplayDriver(int devAdress, int decodeMode, int intensity, int scanLimit)
        //{
        //    if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
        //        return;

        //    try
        //    {
        //        ImportExport.LogMessage("CmdMax7219DisplayInit(Modul: " + devAdress.ToString("X2") + ", decodeMode:" + decodeMode.ToString("X2") + ", intensity: "
        //            + intensity.ToString("X2") + ", scanLimit: " + scanLimit.ToString("X2") + ")", true);

        //        arcazeHid.Command.CmdMax7219DisplayInit(devAdress, decodeMode, intensity, scanLimit);
        //    }
        //    catch (Exception e)
        //    {
        //        ImportExport.LogMessage("CmdMax7219DisplayInit(Modul: " + devAdress.ToString("X2") + ", decodeMode:" + decodeMode.ToString("X2") + ", intensity: "
        //            + intensity.ToString("X2") + ", scanLimit: " + scanLimit.ToString("X2") + ") .. " + e.ToString(), true);
        //    }
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="devAdress">0 .. 15</param>
        /// <param name="digit">SevenSegment Value</param>
        /// <param name="digitMask">0x00 .. 0xFF</param>
        //private void WriteDigitsToDisplayDriver(int devAdress, ref string[] digit, int digitMask)
        //{
        //    Digits = new List<byte>(8);

        //    Digits.Add(byte.Parse(digit[0], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[1], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[2], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[3], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[4], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[5], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[6], NumberStyles.HexNumber));
        //    Digits.Add(byte.Parse(digit[7], NumberStyles.HexNumber));

        //    try
        //    {
        //        if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
        //            return;

        //        arcazeHid.Command.CmdMax7219WriteDigits(devAdress, Digits, digitMask);

        //        if (checkBoxLog.Checked)
        //        {
        //            digitsValue = "";

        //            for (int n = 7; n > -1; n--)
        //            {
        //                digitsValue += Digits[n].ToString("X2") + " ";
        //            }
        //            if (checkBoxLog.Checked)
        //                ImportExport.LogMessage("CmdMax7219WriteDigits(Modul: " + devAdress.ToString("X2") + ", Digits: " + digitsValue + ", Mask: " + (digitMask).ToString("X2") + ")", true);
        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        String digitsValue = "";

        //        for (int n = 0; n < 8; n++)
        //        {
        //            digitsValue = Digits[n].ToString() + " ";
        //        }
        //        ImportExport.LogMessage("CmdMax7219WriteDigits(Modul: " + devAdress.ToString("X2") + ", Digits: " + digitsValue + ", Mask: " + (digitMask).ToString("X2")
        //             + ") .. " + e.ToString(), true);
        //    }
        //}

        #endregion

        #region LED / ADC / Encoder

        private int GetConnectorNumber(int modulNumber, int connector)
        {
            if (modulNumber == 0)
                return (connector == 0 ? 0 : 1);
            else
                return ((connector - 2) + (3 * (modulNumber - 1)));
        }

        private void LEDWriteOutputOld(int moduleNum, int connectorNum, int portNum, int resolution, Double data, int type, bool reverse)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                return;

            if (type != 3)
            {
                if (data < 0.485 && data > 0.001)
                    return;
                if (data > 0.515 && data < 0.999)
                    return;

                data = (data > 0.5 ? 1 : 0); // Arcaze and LED-Driver 2

                if (reverse)
                    data = (data == 1 ? 0 : 1);
            }
            else
            {
                if (reverse)
                    data = 1 - data;

                resolutionValue = Convert.ToUInt32(ResolutionValue(resolution)); // LED-Driver 3
                data *= resolutionValue;

                if (data > resolutionValue)
                    data = resolutionValue;

                if (data < 0)
                    data = 0;
            }

            if (connectorNum > 1)
                connectorNum -= 2;

            //connectorNum = GetConnectorNumber(moduleNum, connectorNum);

            try
            {
                if (checkBoxLog.Checked)
                    ImportExport.LogMessage("WriteOutputPort(Modul: " + moduleNum.ToString("X2") + ", Connector: " + connectorNum.ToString("X2") + ", Pin: " + (portNum + 1).ToString("D2") + ", Value: " + (data == 0 ? "Off" : (type != 3 ? "On" : data.ToString())) + ")", true);

                this.arcazeHid.Command.WriteOutputPort(moduleNum, connectorNum, portNum, Convert.ToUInt32(data), ArcazeCommand.OutputOperators.PlainWrite, false);
                this.arcazeHid.Command.UpdateOutputPorts();
            }
            catch (Exception e)
            {
                ImportExport.LogMessage("WriteOutputPort(Modul: " + moduleNum.ToString("X2") + ", Connector: " + connectorNum.ToString("X2") + ", Pin: " + (portNum + 1).ToString("D2") + ", Value: " + (data == 0 ? "Off" : (moduleNum == 0 ? "On" : data.ToString())) + ") ... " + e.ToString(), true);
            }
        }

        private int ReadADC(int channel)
        {
            int[] inputs = new int[channel];

            try
            {
                if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                    return 0;

                arcazeHid.Command.CmdPrintADC();
                inputs = arcazeHid.Command.CmdReadADC();

                if (checkBoxLog.Checked)
                    ImportExport.LogMessage("ReadADC(" + channel.ToString("D2") + "), Value=" + inputs[channel].ToString("X4"), true);

                return inputs[channel];
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ReadADC(" + channel + ") ... " + f.ToString(), true);
                return 0;
            }
        }

        private int ReadEncoder(int encoderNumber, ref int encoderOldValue)
        {
            GetPortPinOfEncoder(ref encoderNumber, ref port, ref pin);

            //ReadEncoderPinsValues(port, pin, ref encoderPinAvalue, ref encoderPinBvalue);
            //encoderOldValue = (encoderPinAvalue * 2) + encoderPinBvalue;
            //StopWatch.NOP(100); // 100 Microns

            ReadEncoderPinsValues(port, pin, ref encoderPinAvalue, ref encoderPinBvalue);

            encoderNewValue = (encoderPinAvalue * 2) + encoderPinBvalue;

            if (encoderNewValue == encoderOldValue)
                return 0;

            encoderOldValue = 3;

            encoderReturnValue = RotaryEncoder.Encoder(ref encoderNewValue, ref encoderOldValue);
            encoderOldValue = encoderNewValue;

            return encoderReturnValue;
        }

        /// <summary>
        /// Each encoder is represented by a 16 bit signed integer.
        /// The values are absolute, they increment infintely and are never reset. 
        /// They overflow from 32767 to -32786 and vice versa.
        /// </summary>
        //private void ReadEncoderAbsolute()
        //{
        //    if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
        //        return;

        //    try
        //    {
        //        arcazeHid.Command.CmdReadQuadratureEncodersAbsolute(ref encoderValues);
        //    }
        //    catch (Exception f)
        //    {
        //        ImportExport.LogMessage("Read Encoders Absolute ... " + f.ToString(), true);
        //    }
        //}

        /// <summary>
        /// Each encoder is represented by a 16 bit signed integer.
        /// The values are the relative change since the last call of this function. 
        /// On the first call always 0 is returned for all encoders, because there is no reference value yet.
        /// </summary>
        //private void ReadEncoderRelative()
        //{
        //    if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
        //        return;

        //    try
        //    {
        //        arcazeHid.Command.CmdReadQuadratureEncodersRelative(ref encoderValues);
        //    }
        //    catch (Exception f)
        //    {
        //        ImportExport.LogMessage("Read Encoders Relative ... " + f.ToString(), true);
        //    }
        //}

        private void ReadEncoderPinsValues(int port, int pin, ref int pinAvalue, ref int pinBvalue)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                return;

            try
            {
                portValue = arcazeHid.Command.CmdReadPort(port);

                pinAvalue = portValue >> pin;
                pattern = Convert.ToString(pinAvalue, 2);
                pinAvalue = Convert.ToInt32(pattern.Substring(pattern.Length - 1));

                pinBvalue = portValue >> (pin + 1);
                pattern = Convert.ToString(pinBvalue, 2);
                pinBvalue = Convert.ToInt32(pattern.Substring(pattern.Length - 1));
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ReadPort(" + port + ") .. " + f.ToString(), true);
            }
        }

        private void ReadPinValue(int port, int pin, ref int value, ref bool reverse)
        {
            try
            {
                if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                    return;

                value = arcazeHid.Command.CmdReadPort(port);

                value = value >> pin;
                pattern = Convert.ToString(value, 2);
                value = Convert.ToInt32(pattern.Substring(pattern.Length - 1));

                if (!reverse)
                    value = (value == 1 ? 0 : 1); // Invert it
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ReadPort(" + port + ") .. " + f.ToString(), true);
            }
        }

        /// <summary>
        /// Set the pin direction on the arcaze
        /// </summary>
        /// <param name="port">Connector A = 0;B = 1;C = 2</param>
        /// <param name="pin">0 - 19</param>
        /// <param name="direction">0 = input; 1 = output</param>
        private void SetPinDirection(int port, int pin, int direction)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                return;

            try
            {
                ImportExport.LogMessage("SetPinDirection(Connector: " + port.ToString("X2") + ", Pin: " + (pin + 1).ToString("D2") + ", Direction: " + (direction == 1 ? "Output" : "Input") + ")", true);

                arcazeHid.Command.CmdSetPinDirection(port, pin, direction);
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("SetPinDirection(Connector: " + port.ToString("X2") + ", Pin: " + (pin + 1).ToString("D2") + ", Direction: " + (direction == 1 ? "Output" : "Input") + ") .. " + f.ToString(), true);
            }
        }

        #endregion

        #region events

        //private void DeviceRemoved(object sender, HidEventArgs hidEventsArgs)
        //{
        //    this.arcazeHid.Disconnect();
        //    arcazeFound = false;
        //    timerstate = State.startup;
        //}

        //private void OurDeviceRemoved(object sender, HidEventArgs hidEventsArgs)
        //{
        //    this.arcazeHid.Disconnect();
        //    arcazeFound = false;
        //    timerstate = State.startup;
        //}

        //private void DeviceReceived(object sender, HidEventArgs hidEventsArgs)
        //{
        //    //this.arcazeHid.Connect(hidEventsArgs.DeviceInfo.Path);
        //    arcazeFound = false;
        //    timerstate = State.startup;
        //}

        //private void OurDeviceReceived(object sender, HidEventArgs hidEventsArgs)
        //{
        //    arcazeFound = false;
        //    timerstate = State.startup;
        //}

        private void ArcazeChanged(object sender, EventArgs e)
        {
            ChangeArcase(presentArcaze[ComboBoxArcaze.SelectedIndex].Path);
        }

        private void ButtonCopyToClipboard_Click(object sender, EventArgs e)
        {
            CopyListBoxToClipboard(listBox1);
        }

        private void ButtonFind_Click(object sender, EventArgs e)
        {
            InitAll();
        }

        private void ButtonInit_Click(object sender, EventArgs e)
        {
            buttonInit.Visible = false;
            InitAll();
        }

        private void ButtonLogClear_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void ButtonOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialog1.Filter = "D.A.C. Data Files *.xml|*.xml|*.*|*.*";
            openFileDialog1.Title = "Select a .xml File";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SwitchInverseOff();

                textBoxLastFile.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\', openFileDialog1.FileName.Length - 1) + 1);
                ImportExport.LogMessage("Used configuration file " + textBoxLastFile.Text, true);

                dataSetFilename = openFileDialog1.FileName;

                dataSetDisplaysLEDs.Clear();
                ImportExport.XmlToDataSet(openFileDialog1.FileName, dataSetDisplaysLEDs);
                ImportExport.DatasetToXml("config.xml", dataSetConfig);

                InitTables();
                InitAll();
                this.Text = this.Text.Substring(0, (this.Text.IndexOf("(") - 1)) + "  (Config. with " + textBoxLastFile.Text + ")";
            }
        }

        private void ButtonPackage_Click(object sender, EventArgs e)
        {
            if (textBoxTestDataPackage.Text.Trim().Length > 0)
                receivedData = textBoxTestDataPackage.Text;
        }

        private void ButtonReadEncoderAbs_Click(object sender, EventArgs e)
        {
            //ReadEncoderAbsolute();
            //SetEnconderValue();
        }

        private void ButtonReadEncoderRel_Click(object sender, EventArgs e)
        {
            //ReadEncoderRelative();
            //SetEnconderValue();
        }

        private void ButtonReset_Click(object sender, EventArgs e)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected)
                return;

            timerstate = State.reset;
        }

        private void ButtonSaveConfig_Click(object sender, EventArgs e)
        {
            SaveConfig();
            labelPleaseWait.Text = "Configuation saved.";
        }

        private void ButtonSaveToFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            saveFileDialog1.Filter = "D.A.C. Data Files *.xml|*.xml";
            saveFileDialog1.Title = "Save a .xml File";
            saveFileDialog1.FileName = textBoxLastFile.Text;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxLastFile.Text = saveFileDialog1.FileName.Substring(saveFileDialog1.FileName.LastIndexOf('\\', saveFileDialog1.FileName.Length - 1) + 1);
                this.Text = this.Text.Substring(0, (this.Text.IndexOf("(") - 1)) + "  (Config. with " + textBoxLastFile.Text + ")";

                dataSetDisplaysLEDs.AcceptChanges();

                ImportExport.DatasetToXml(saveFileDialog1.FileName, dataSetDisplaysLEDs);
                SaveConfig();

                labelPleaseWait.Text = "All data saved.";
            }
        }

        private void ButtonSendPin_Click(object sender, EventArgs e)
        {
            arcazeDeviceIndex = GetArcazeInstanceFromLEDs(ComboBoxArcaze.Text);

            if (arcazeDeviceIndex > -1)
            {
                if (InitLEDDriver())
                {
                    switch (comboBoxModuleType.Text)
                    {
                        case "Display-Driver":
                            break;

                        case "LED-Driver 2":
                            arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(comboBoxModul.SelectedIndex, comboBoxConnector.SelectedIndex,
                                comboBoxPin.SelectedIndex, 1, Convert.ToDouble(checkBoxPinSetValue.Checked ? 1 : 0), 2, false, checkBoxLog.Checked);
                            break;

                        case "LED-Driver 3":
                            if (comboBoxLEDValue.Text == "")
                                comboBoxLEDValue.Text = "1.00";

                            pinString = comboBoxLEDValue.Text.Replace(",", ".");

                            arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(comboBoxModul.SelectedIndex, comboBoxConnector.SelectedIndex,
                                comboBoxPin.SelectedIndex, 8, double.Parse(pinString, CultureInfo.InvariantCulture), 3, false, checkBoxLog.Checked);
                            break;

                        case "Arcaze USB":
                            arcazeDevice[arcazeDeviceIndex].SetPinDirection(comboBoxConnector.SelectedIndex, comboBoxPin.SelectedIndex, 1); // direction 1 -> output

                            arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(comboBoxModul.SelectedIndex, comboBoxConnector.SelectedIndex,
                                comboBoxPin.SelectedIndex, 1, Convert.ToDouble(checkBoxPinSetValue.Checked ? 1 : 0), 4, false, checkBoxLog.Checked);
                            break;
                    }
                }
            }
        }

        private void ButtonSendTestPattern_Click(object sender, EventArgs e)
        {
            SystemCheck();
        }

        private void ButtonStop_Click(object sender, EventArgs e)
        {
            stop = !stop;
            buttonStop.Text = (stop ? "Start" : "Stop");
            buttonStop.ForeColor = (stop ? Color.Green : Color.Red);

            if (stop)
                ImportExport.LogMessage("Package processing stopped .. ", true);
            else
                ImportExport.LogMessage("Package processing restarted .. ", true);
        }

        private void ButtonWriteDigits_Click(object sender, EventArgs e)
        {
            ActivateArcaze(ComboBoxArcaze.Text, false);
            arcazeDeviceIndex = GetArcazeInstanceFromDisplays(ComboBoxArcaze.Text);

            if (arcazeDeviceIndex > -1)
            {
                digitString = "";

                for (int n = 0; n < textBoxDezValue.Text.Length; n++)
                {
                    characterValue = (byte)Convert.ToChar(textBoxDezValue.Text.Substring(n, 1));

                    if (characterValue >= 44 && characterValue <= 57 && characterValue != 47)
                        digitString += textBoxDezValue.Text.Substring(n, 1);
                }
                textBoxDezValue.Text = digitString;

                ConvertDezToSevenSegment(ref digit);

                arcazeDevice[arcazeDeviceIndex].InitDisplayDriver(int.Parse(comboBoxDevAddress.Text, NumberStyles.HexNumber), 0, trackBarDisplayBrightness.Value, 8);

                arcazeDevice[arcazeDeviceIndex].WriteDigitsToDisplayDriver(int.Parse(comboBoxDevAddress.Text, NumberStyles.HexNumber), ref digit,
                    int.Parse(textBoxDigitMask.Text, NumberStyles.HexNumber), checkBoxLog.Checked, checkBoxReverse.Checked, segmentIndex,
                    Convert.ToInt32(cbRefreshCycle.Text), Convert.ToInt32(cbDelayRefresh.Text));
            }
        }

        private void CheckBoxPinSetValue_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxPinSetValue.Text = (checkBoxPinSetValue.Checked ? "LED On" : "LED Off");
        }

        private void CheckBoxUseLeftPadding_CheckedChanged(object sender, EventArgs e)
        {
            useLeftPadding = checkBoxUseLeftPadding.Checked;
        }

        private void ComboBoxLEDModul_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxModul.SelectedIndex == 0)
                comboBoxConnector.SelectedIndex = 0;
            else
                comboBoxConnector.SelectedIndex = 2;
        }

        private void ComboBoxArcaze_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ComboBoxArcaze.Text != "")
                ActivateArcaze(ComboBoxArcaze.Text, false);
        }

        private void ComboBoxModuleType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxModuleType.Text)
            {
                case "Display-Driver": // Display-Driver
                    break;
                case "LED-Driver 2":
                    if (comboBoxModul.SelectedIndex == 0)
                        comboBoxModul.SelectedIndex = 1;
                    if (comboBoxConnector.SelectedIndex == 0)
                        comboBoxConnector.SelectedIndex = 2;

                    labelTestLEDValue.Visible = false;
                    comboBoxLEDValue.Visible = false;
                    checkBoxPinSetValue.Visible = true;

                    InitLEDDriver();

                    break;
                case "LED-Driver 3":
                    if (comboBoxModul.SelectedIndex == 0)
                        comboBoxModul.SelectedIndex = 1;
                    if (comboBoxConnector.SelectedIndex == 0)
                        comboBoxConnector.SelectedIndex = 2;

                    labelTestLEDValue.Visible = true;
                    comboBoxLEDValue.Visible = true;
                    checkBoxPinSetValue.Visible = false;

                    InitLEDDriver();

                    break;
                case "Arcaze USB":
                    if (comboBoxModul.SelectedIndex > 0)
                        comboBoxModul.SelectedIndex = 0;
                    if (comboBoxConnector.SelectedIndex > 1)
                        comboBoxConnector.SelectedIndex = 0;

                    labelTestLEDValue.Visible = false;
                    comboBoxLEDValue.Visible = false;
                    checkBoxPinSetValue.Visible = true;

                    break;
            }
        }

        private void dataGridViewArcaze_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewClickable_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridViewRow row = dataGridViewClickable.Rows[e.RowIndex];
            //row.Cells[3].Value = "D " + row.Cells[0].Value + " - " + "B " + row.Cells[1].Value + " - " + row.Cells[2].Value;
        }

        private void DataGridViewClickable_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewDisplays_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!arcazeFound)
                return;

            DataGridViewRow row = dataGridViewDisplays.Rows[e.RowIndex];

            if (e.ColumnIndex > 1 && e.ColumnIndex < 7 && e.ColumnIndex != 3)
                buttonInit.Visible = true;

            if (e.ColumnIndex == 3)
            {
                dcsExportID = row.Cells["dcsExportIDDisplays"].Value.ToString();
                checkTest = Convert.ToBoolean(row.Cells["displayTest"].Value);
                receivedData = ":" + dcsExportID + "=" + (checkTest ? "88888888" : "-") + ":";
            }
        }

        private void DataGridViewDisplays_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //ImportExport.LogMessage("DataGridViewDisplays : " + e.ToString(), true);
        }

        private void DataGridViewJoystick_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //ImportExport.LogMessage("DataGridViewJoystick : " + e.ToString(), true);
        }

        private void DataGridViewEncoderValues_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridViewEncoderValues.Rows[e.RowIndex];

            if (row.Cells["Arcaze"].Value != null && row.Cells["EncoderNumber"].Value != null)
                row.Cells["ArcazeAndEncoder"].Value = row.Cells["Arcaze"].Value.ToString() + " - " + row.Cells["EncoderNumber"].Value.ToString();
            //dataGridViewEncoderSend.Refresh();
        }

        private void dataGridViewEncoderValues_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 2 && e.ColumnIndex < 7 && arcazeFound)
                buttonInit.Visible = true;
        }

        private void dataGridViewEncoderValues_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewExportID_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridViewExportID.Rows[e.RowIndex];
            row.Cells[2].Value = row.Cells[0].Value + " - " + row.Cells[1].Value;
        }

        private void DataGridViewLEDs_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!arcazeFound)
                return;

            DataGridViewRow row = dataGridViewLEDs.Rows[e.RowIndex];

            if (e.ColumnIndex > 3 && e.ColumnIndex < 11)
                buttonInit.Visible = true;

            if (e.ColumnIndex == 3) // Send test package
            {
                dcsExportID = row.Cells["dcsExportIDLEDs"].Value.ToString();
                resolution = Convert.ToInt32(row.Cells["ledResolution"].Value);
                checkTest = Convert.ToBoolean(row.Cells["ledsTest"].Value);

                switch (resolution)
                {
                    case 1:
                        receivedData = ":" + dcsExportID + "=" + (checkTest ? "1" : "0") + ":";
                        break;
                    case 2:
                        receivedData = ":" + dcsExportID + "=" + (checkTest ? "3" : "0") + ":";
                        break;
                    case 4:
                        receivedData = ":" + dcsExportID + "=" + (checkTest ? "15" : "0") + ":";
                        break;
                    case 8:
                        receivedData = ":" + dcsExportID + "=" + (checkTest ? "255" : "0") + ":";
                        break;
                }
            }

            if (e.ColumnIndex == 6)
            {
                if (row.Cells["Modul"].Value.ToString() == "0" && row.Cells["modulTypeID"].Value.ToString() != "Arcaze USB")
                {
                    row.Cells["Modul"].Value = 1;
                    row.Cells["portLEDs"].Value = "Port C";
                }

                if (row.Cells["Modul"].Value.ToString() != "0" && row.Cells[6].Value.ToString() == "Arcaze USB")
                {
                    row.Cells["Modul"].Value = 0;
                    row.Cells["portLEDs"].Value = "Port A";
                }
            }
        }

        private void DataGridViewLEDs_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //ImportExport.LogMessage("DataGridViewLEDs : " + e.ToString(), true);
        }

        private void dataGridViewKeystrokes_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 1 && e.ColumnIndex < 6 && arcazeFound)
                buttonInit.Visible = true;
        }

        private void dataGridViewKeystrokes_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewSwitches_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 1 && e.ColumnIndex < 9 && arcazeFound)
                buttonInit.Visible = true;
        }

        private void DataGridViewSwitches_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //ImportExport.LogMessage("dataGridViewSwitches : " + e.ToString(), true);
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (udpThread != null)
            {
                UDP.ListenerClose();
                udpThread.Abort();
            }
            timerMain.Stop();

            try
            {
                rows = new DataRow[] { };
                rows = dataSetConfig.Tables["SendKeysTimeBased"].Select("Active=" + true + " AND Sended=" + false + " AND State = 'On exit'");

                SendKeystrokesTimeBased(ref rows);

                SwitchAll(off);
                SwitchInverseOff();
                DisconnectArcaze();
                //Environment.Exit(0);
            }
            catch { }
        }

        //private void LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        //{
        //    if (linkLabel1.Text.Substring(0, 4).ToUpper() == "HTTP")
        //    {
        //        try
        //        {
        //            Process.Start(linkLabel1.Text);
        //        }
        //        catch (Exception f)
        //        {
        //            ImportExport.LogMessage("LinkLabel_LinkClicked ... " + f.ToString(), true);
        //        }
        //    }
        //}

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetLogs();
            if (tabControl1.SelectedIndex > 4)
            {
                if (buttonInit.Visible)
                {
                    buttonInit.Visible = false;
                    InitAll();
                }
            }
            if (tabControl1.SelectedIndex == 5)
            {
                ActivateArcaze(ComboBoxArcaze.Text, false);
                labelOnOff.Text = (ledOn ? "Switch 'Off'" : "Switch 'On'");
            }
        }

        private void TextBoxIntervalTimer_TextChanged(object sender, EventArgs e)
        {
            timerInterval = Convert.ToInt32(textBoxIntervalTimer.Text);
        }

        private void TrackBarDisplayBrightness_Scroll(object sender, EventArgs e)
        {
            InitDisplaysDrivers();
        }

        private void TrackBarLEDDriverBrightness_Scroll(object sender, EventArgs e)
        {
            SetPinDirectionForLEDs();
        }

        #endregion

        private void dataGridViewDisplays_ColumnDividerDoubleClick(object sender, DataGridViewColumnDividerDoubleClickEventArgs e)
        {

        }
    }
}
