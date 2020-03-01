
#region Lizenz
/*
        ----------------------------------------------------------------------
         Copyright (c) 2015 Heinz-Joerg Puhlmann

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

using SimpleSolutions.Usb;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using WindowsInput;
using Newtonsoft.Json;

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
		         Added checkbox 'R => Reverse' for displays.
		0.952    Bug fixes;                                                                                        12.07.2014
		0.953    Inverse logic for switches. Added column 'R -> Reverse' for switches and keystrokes.              17.07.2014
		         Multi position switches are working now.
		         Tab 'Encoders' is now active. Added logic after changed 'Module Type' in tab 'LEDs'.
		         Added checkbox 'T -> Test' in tab 'LEDs'.
		0.954    Bug fixes;                                                                                        18.07.2014
		0.955    Optimizing the layout; Added column token to tab 'help'; Added 'T -> Test' in tab 'Displays';     21.07.2014
		         Cleanup tab 'Help'; Added 'R -> Reverse' in tab 'LEDs';
		0.956    Bug fixes; Added 'S -> Send not zero value' in tab 'Switches'; tab 'help' changed;                22.07.2014
		0.957    Bug fixes with the focus on encoders;                                                             23.07.2014
		0.958    Resource optimization (RAM and CPU); Bug fixes;                                                   24.07.2014
		0.959    Instantiate the arcaze for LED output; Bugfixes encoder for connector B;                          03.08.2014
		0.960    Instantiate the arcaze for display output;                                                        04.08.2014
		0.961    Bug fixes tab 'Test'; Bug fixes instantiate arcaze for LED output and display output;             05.08.2014
		0.962    Bug fixes tab 'Test';                                                                             06.08.2014
		0.963    Bug fixes; Tab 'Keystrokes' enabled;                                                              06.08.2014
		0.964    Bug fixes tab 'LEDs'; Tab 'Test' repaired;                                                        07.08.2014
		0.965    Optimize data transfere; Fix 'Display - Driver 32' problems;                                      08.08.2014
		0.966    Display refresh after 1 ms;                                                                       18.08.2014
		0.967    Bugfixes 'left padding', lost of IP address after save;                                           19.08.2014
		0.968    Bugfixes dezimal points; for Displays 2 times refresh after 0.3 ms;                               29.08.2014
		0.969    Added in tab 'Configuration' delay and refresh cycle for display driver 32;                       31.08.2014
		0.970    Bugfixes 'Export ID' for small IDs like '50'.                                                     29.09.2014
		0.971    Optimize logging                                                                                  02.11.2014
		0.972    Optimize listener                                                                                 03.11.2014
		0.973    Optimize listener: Now it is possible to receive multi packages;                                  04.11.2014
		0.974    Package analyse: Now I use LastIndexOf for a ID;                                                  15.11.2014
		0.975    With firmware update 5.64 we get functional encoders;                                             01.04.2015
		0.976    New section in Tab 'Test' for ADC test;                                                           02.04.2015
		0.977    Bugfix ADC: 'arcazeHid.Command.CmdPrintADC()' is not more implemented.                            22.04.2015
		0.978    Optimize Tab ADC.                                                                                 21.05.2015
		0.979    Added field 'type' in table clickable and export ID.                                              22.05.2015
		0.980    Implement selected pulldown for field '.. Desciption..' and 'Resource ...' in tab 'Display',      23.05.2015
		         'LEDs', 'Encoders' and 'ADC' 
		0.982    Optimize database restauration                                                                    25.05.2015
		0.983    Added checkbox 'R => Reverse' for Encoders.                                                       01.07.2015
		---------------------------------------------------------------------------------------------------------------------
     */

    public partial class FormMain : Form
    {
        #region member

        enum State
        {
            sendJson,
            initConfig,
            sendConfig,
            sendKeys,
            readNewFile,
            startup,
            init,
            systemCheck,
            checkLEDS,
            checkDisplays,
            run,
            reset,
            stop
        }

        private string temp = "";
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

        public static DataSet dsJSON = new DataSet();
        public static DataSet dsJson;
        public static DataTable dtjsonLamps;
        public static DataTable dtjsonSwitches;
        public static DataTable dtJson;

        public static string json = "";

        private static DataRow[] dataRows = new DataRow[] { };
        public static DataRow dataRow = null;

        public static DataTable dtLamps;
        public static DataTable dtDisplays;
        public static DataTable dtDcs_ID;


        bool[] checkBoxDP = new bool[8] { false, false, false, false, false, false, false, false };
        bool[] checkBox = new bool[8];

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
        private bool reverseEncoder = false;

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

        private double minValue_ADC = 0.0;
        private double maxValue_ADC = 0.0;
        private double m_ADC = 0.0;
        private double n_ADC = 0.0;
        private double pinValueMax = 1.0;

        private int encoderDeltaValue = 0;
        private int encoderIdentity = 0;

        private int encoderNewValue = 0;
        private int encoderOldValue = 0;
        private int encoderPinAvalue = 0;
        private int encoderPinBvalue = 0;
        private int encoderReturnValue = 0;

        private int encoderNo = 0;
        private int cannelADC = 0;
        private int adcOldValue = 0;
        private string encoderName = "";
        private int mainTestLoop = 2;
        private int module = 0;
        private int numberInputChanged = 0;
        private int loopCounter = 0;
        private int loopMax = 50;
        private int pin = 0;
        private int pinVal = 0;
        private int port = 0;
        private int portValue = 0;

        private int resolution = 1;
        private int segmentIndex = 0;
        private int testLoop = 0;
        private int timerInterval = 20;
        private int timerIntervalTest = 100;
        private int timerIntervalReset = 5000;
        private int type = 0;
        private int connectCounter = 40;
        private int connectCounterMax = 20;
        private int configID = -1;
        private string[] receivedItems = new string[] { };

        private double encoderLastValue = 0;

        private string[] digit = new string[8] { "", "", "", "", "", "", "", "" };
        private string arcazeAddress = "";
        private string arcazeFromGrid = "";
        private string dataSetFilename = "\\Dataset.xml";
        private string dcsExportID = "";
        private string devAddress = "";
        private string dezValue = "";
        private string digitMask = "";
        private string digitString = "";
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
        public static string receivedData = "";
        private string receivedDataBackup = "";
        private const string searchStringForFile = "File";
        private string sendKey = "";
        private string optionKeyCode = "";
        private string optionKeyCodeTwo = "";
        private string valueON = "1.0";
        private string valueOff = "0.0";
        private string pinMax = "1.0";
        private string arcaze = "";
        private static string start = "=" + '"';
        public static bool logDetail = false;
        private bool idFound = false;

        private static bool debug = false;

        #endregion

        public FormMain()
        {
            InitializeComponent();

            try
            {
                ImportExport.LogMessage("Application started ...", true);
                ImportExport.LogMessage(labelVersion.Text, true);

                try
                {
                    textBoxLastFile.Text = "A-10C.xml";
                    checkBoxWriteLogsToHD.Checked = false;
                    checkBoxLog.Checked = false;
                    logDetail = checkBoxLog.Checked;
                    textBoxIP.Text = "127.0.0.1";
                    textBoxPortListener.Text = "26026";
                    textBoxPortSender.Text = "26027";
                    textBoxTestDataPackage.Text = "";
                    textBoxIntervalTimer.Text = "5";

                    ImportExport.XmlToDataSet(Application.StartupPath + "\\" + "config.xml", dataSetConfig);
                    ImportExport.LogMessage("Loaded config.xml ... ", true);

                    checkBoxLog.Checked = false; // Convert.ToBoolean(dataSetConfig.Tables["Config"].Rows[0]["LogAllActions"]);

                    if (checkBoxLog.Checked)
                    {
                        logDetail = true;
                    }

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
                    {
                        textBoxLastFile.Text += ".xml";
                    }

                    ImportExport.XmlToDataSet(Application.StartupPath + "\\" + textBoxLastFile.Text, dataSetDisplaysLEDs);
                    readFile = textBoxLastFile.Text;
                    lastFile = readFile;

                    textBoxIntervalTimer.Text = dataSetConfig.Tables["Config"].Rows[0]["IntervalTimer"].ToString();
                    ImportExport.LogMessage("Interval timer: " + textBoxIntervalTimer.Text + "[ms]", true);
                }
                catch (Exception f) { ImportExport.LogMessage("Loading error: config.xml ... " + f.ToString(), true); }

                if (File.Exists(logFile))
                {
                    try { File.Delete(logFile); }
                    catch (Exception f) { ImportExport.LogMessage("File.Delete ... " + f.ToString(), true); }
                }

                try
                {
                    InitTables();

                    tabControl1.TabPages.Remove(Info);
                    tabControl1.TabPages.Remove(Experimental);
                    tabControl1.TabPages.Remove(Copyright);
                    tabControl1.TabPages.Remove(Keystrokes);
                    tabControl1.SelectedIndex = 5;

                    RefreshPulldown();
                    StartTimer();
                    stop = false;
                }
                catch (Exception f) { ImportExport.LogMessage("Application start problem ... " + f.ToString(), true); }
            }
            catch (Exception f) { ImportExport.LogMessage("Application start problem ... " + f.ToString(), true); }
        }

        private void TimerMain_Tick(object sender, EventArgs e)
        {
            #region send keys timebased

            if (lStateEnabled && !stop)
            {
                lStateEnabled = false;

                try
                {
                    if (sendOpenKeysTimeBased)
                    {
                        rows = new DataRow[] { };
                        rows = dataSetConfig.Tables["SendKeysTimeBased"].Select("Active=" + true + " AND Sended=" + false + " AND State = 'On startup'");

                        SendKeystrokesTimeBased(ref rows);
                    }
                }
                catch (Exception f) { ImportExport.LogMessage("Send keys timebased: ... " + f.ToString(), true); }

                lStateEnabled = true;
            }

            #endregion

            #region start listener

            if (UDP.CloseListener && textBoxPortListener.Text.Trim().Length >= 4)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
                        if (udpThread != null)
                            udpThread.Abort();

                        udpThread = new Thread(new ThreadStart(StartListener))
                        {
                            IsBackground = true
                        };
                        udpThread.Start();

                        this.Text = this.Text + "  (Config. with " + textBoxLastFile.Text + ")";
                    }
                    catch (Exception f) { ImportExport.LogMessage("StartListener: ... " + f.ToString(), true); }

                    timerstate = State.startup;

                    lStateEnabled = true;
                }
            }

            #endregion

            if (timerstate == State.readNewFile)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;
                    stop = true;

                    try
                    {
                        if (readFile.Length > 0 && readFile != lastFile)
                        {
                            if (!File.Exists(Application.StartupPath + "\\" + readFile))
                            {
                                ImportExport.LogMessage("File not found: " + readFile + " ... ", true);
                            }
                            else
                            {
                                SwitchAll(off);
                                SwitchInverseOff();

                                ImportExport.LogMessage("Read configuration file : " + readFile, true);
                                ImportExport.XmlToDataSet(Application.StartupPath + "\\" + readFile, dataSetDisplaysLEDs);
                                ImportExport.LogMessage("Used configuration file : " + readFile , true);

                                FindAllArcaze();
                                RefreshPulldown();
                                InitDrivers();

                                lastFile = readFile;
                                textBoxLastFile.Text = lastFile;
                                this.Text = "D.A.C. - DCS Arcaze Communicator  (Config. with " + textBoxLastFile.Text + ")";

                                timerMain.Interval = timerInterval;
                                receivedData = receivedDataBackup;
                            }
                        }
                    }
                    catch (Exception f) { ImportExport.LogMessage("State readfile: " + readFile + " ... " + f.ToString(), true); }

                    loadFile = false;

                    if (arcazeFound || debug)
                    {
                        timerstate = State.sendConfig;
                    }
                    else
                    {
                        timerstate = State.startup;
                    }
                    stop = false;
                    lStateEnabled = true;
                }
            }

            if (timerstate == State.startup)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
                        labelPleaseWait.Text = "Startup ... ";
                        timerMain.Interval = timerInterval;

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

                            labelPleaseWait.Text = "";
                            timerMain.Interval = timerInterval;
                            timerstate = State.init;
                        }
                    }
                    catch (Exception f) { ImportExport.LogMessage("State Startup: ... " + f.ToString(), true); }

                    lStateEnabled = true;
                }
            }

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

            if (timerstate == State.init)
            {
                if (arcazeFound)
                {
                    if (lStateEnabled)
                    {
                        lStateEnabled = false;

                        try
                        {
                            labelPleaseWait.Text = "Init is running. Please wait ..";

                            InitDrivers();

                            SwitchAll(on);
                            SwitchAll(off);

                            testLoop = 0;
                            mainTestLoop = 0;
                            timerstate = State.systemCheck;
                        }
                        catch (Exception f) { ImportExport.LogMessage("State Init: ... " + f.ToString(), true); }

                        lStateEnabled = true;
                    }
                }
            }

            if (timerstate == State.checkLEDS)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
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
                    }
                    catch (Exception f) { ImportExport.LogMessage("State check LEDs ... " + f.ToString(), true); }

                    lStateEnabled = true;
                }
            }

            if (timerstate == State.checkDisplays)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
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

                                        pattern = ":" + dcsExportID + "=" + (ledOn ? "88888888" : "") + ":";
                                        labelPleaseWait.Text = "Display Test " + (ledOn ? "On" : "Off");

                                        receivedData = pattern;
                                    }
                                    testLoop++;
                                }
                                catch (Exception f) { ImportExport.LogMessage("Timerstate check Displays ... " + f.ToString(), true); }
                            }
                        }
                        else
                        {
                            labelPleaseWait.Text = "";
                            buttonSendTestPattern.Enabled = true;
                            timerstate = State.systemCheck;
                        }
                    }
                    catch (Exception f) { ImportExport.LogMessage("State check Displays ... " + f.ToString(), true); }

                    lStateEnabled = true;
                }
            }

            if (timerstate == State.systemCheck)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
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
                            comboBoxDevAddress.SelectedIndex = 0;
                            comboBoxConnector.SelectedIndex = 0;
                            comboBoxPin.SelectedIndex = 0;
                            comboBoxModuleType.SelectedIndex = 3;
                            comboBoxModul.SelectedIndex = 0;
                        }

                        labelPleaseWait.Text = "";
                        buttonInit.Visible = false;
                        timerMain.Interval = timerInterval;

                        timerstate = State.sendConfig;
                    }
                    catch (Exception f) { ImportExport.LogMessage("State system check ... " + f.ToString(), true); }

                    lStateEnabled = true;
                }
            }

            if (timerstate == State.initConfig)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
                        package = "{'Registration': {'Name': 'DAC', 'IP': '" + textBoxIP.Text.Trim() + "', 'Port': '" + textBoxPortListener.Text.Trim() + "'}}";
                        package = package.Replace("'", '"'.ToString());
                    }
                    catch { }

                    timerstate = State.sendConfig;

                    lStateEnabled = true;
                }
            }

            if (timerstate == State.sendConfig)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    if (!stop)
                    {
                        try
                        {
                            if (configID == -1)
                            {
                                connectCounter++;

                                if (connectCounter >= connectCounterMax * 10)
                                {
                                    package = "{'Registration': {'Name': 'DAC', 'IP': '" + textBoxIP.Text.Trim() + "', 'Port': '" + textBoxPortListener.Text.Trim() + "'}}";
                                    package = package.Replace("'", '"'.ToString());

                                    UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), package);
                                    connectCounter = 0;
                                    ImportExport.LogMessage("Send connection: " + package, true);
                                }
                            }
                            else
                            {
                                if (!loadFile)
                                {
                                    connectCounter = connectCounterMax / 2;
                                    GenerateJSONDataset();

                                    RefreshAllDatagrids();
                                    InitDrivers();

                                    loopCounter = 0;
                                    timerstate = State.run;
                                }
                            }
                        }
                        catch { }
                    }
                    lStateEnabled = true;
                }
            }

            if (timerstate == State.reset)
            {
                if (lStateEnabled)
                {
                    lStateEnabled = false;

                    try
                    {
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
                    }
                    catch (Exception f) { ImportExport.LogMessage("State Reset: ... " + f.ToString(), true); }

                    lStateEnabled = true;
                }
            }

            #region LED dimming / blinking

            if (timerstate == State.run && !stop && lStateEnabled)
            {
                lStateEnabled = false;

                try
                {
                    for (int n = 0; n < dataGridViewLEDs.Rows.Count; n++)
                    {
                        if (Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["activeLEDs"].Value) &&
                            Convert.ToBoolean(dataGridViewLEDs.Rows[n].Cells["checkBoxInit"].Value))
                        {
                            comboCell = (DataGridViewComboBoxCell)dataGridViewLEDs.Rows[n].Cells["modulTypeID"];

                            if (comboCell.Value == null)
                                type = 4;
                            else
                                type = comboCell.Items.IndexOf(comboCell.Value) + 1;

                            if (type != 1 && Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["dimmMode"].Value) != 0)
                            {
                                if (dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value != DBNull.Value)
                                {
                                    if (double.Parse(dataGridViewLEDs.Rows[n].Cells["valueLEDs"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture) > 0.5)
                                    {
                                        arcazeFromGrid = "";

                                        if (dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value != null) { arcazeFromGrid = dataGridViewLEDs.Rows[n].Cells["arcazeLEDs"].Value.ToString(); }

                                        if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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

                                            if (textBoxIntervalTimer.Text == "") { textBoxIntervalTimer.Text = "20"; }

                                            dimmingQuotient = double.Parse(dataGridViewLEDs.Rows[n].Cells["dimmTime"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);

                                            dimmingInterval = double.Parse(textBoxIntervalTimer.Text.Replace(",", "."), CultureInfo.InvariantCulture) /
                                                             (dimmingQuotient * 250);

                                            if (dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value != DBNull.Value)
                                            {
                                                pinMax = dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value.ToString().Replace(",", ".");
                                            }
                                            else
                                            {
                                                pinMax = "1.0";
                                                dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value = "1.0";
                                            }

                                            if (Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["dimmMode"].Value) == 2) // dimming up
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

                                            if (dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value != DBNull.Value)
                                            {
                                                pinMax = dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value.ToString().Replace(",", ".");
                                            }
                                            else
                                            {
                                                pinMax = "1.0";
                                                dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value = "1.0";
                                            }

                                            switch (type)
                                            {
                                                case 2: // LED Driver 2
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port, ref pin, ref resolution, Convert.ToDouble(dimmingValue > 0.5 ? 1 : 0), type, isReverse, false);
                                                    break;

                                                case 3: // LED Driver 3
                                                    pinValueMax = double.Parse(pinMax, System.Globalization.CultureInfo.InvariantCulture);

                                                    if (pinValueMax > 1.0) { pinValueMax = 1.0; }
                                                    if (pinValueMax < 0.0) { pinValueMax = 0.0; }

                                                    dimmingValue *= pinValueMax;
                                                    dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value = pinValueMax.ToString();

                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port, ref pin, ref resolution, dimmingValue, type, isReverse, false);
                                                    break;

                                                case 4: // Arcase USB
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port, ref pin, ref resolution, Convert.ToDouble(dimmingValue > 0.5 ? 1 : 0), type, isReverse, false);
                                                    break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception f) { ImportExport.LogMessage("State Run: ... " + f.ToString(), true); }

                lStateEnabled = true;
            }

            #endregion

            #region new data

            //if (lStateEnabled && !stop && arcazeFound && receivedData.Length > 4) // New Data ?
            if (lStateEnabled && receivedData.Length > 4) // New Data ?
            {
                lStateEnabled = false;

                if (checkBoxLog.Checked && receivedData.Length > 4)
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
                        if (checkBoxLog.Checked && receivedData.Length > 4)
                            ImportExport.LogMessage("Processing package: " + receivedData, true);

                        receivedData += newline;
                        GrabValues();
                    }
                    //dataSetDisplaysLEDs.AcceptChanges();
                }
                receivedData = "";

                lStateEnabled = true;
            }
            #endregion

            #region new joystick button

            //if (Joystick.joystickActiv && !stop)
            //{
            //    if (lStateEnabled)
            //    {
            //        lStateEnabled = false;

            //        try
            //        {
            //            if (dataSetConfig.Tables["Joysticks"].Rows.Count == 0)
            //            {
            //                for (int n = 0; n < Joystick.joystickName.Count; n++)
            //                {
            //                    row = dataSetConfig.Tables["Joysticks"].NewRow();
            //                    row[0] = Joystick.joystickName[n];
            //                    dataSetConfig.Tables["Joysticks"].Rows.Add(row);
            //                }
            //            }

            //            if (dataSetConfig.Tables["SendKeysFromJoystick"].Rows.Count > 0)
            //            {
            //                for (int n = 0; n < Joystick.joystickName.Count; n++)
            //                {
            //                    Joystick.UpdateJoystick(n, ref buttonsPressed);

            //                    if (buttonsPressed != "")
            //                    {
            //                        if (checkBoxLog.Checked)
            //                            ImportExport.LogMessage("Joystick: " + Joystick.joystickName[n] + " .. processing buttons: --> " + buttonsPressed + " <--", true);

            //                        rows = new DataRow[] { };
            //                        rows = dataSetConfig.Tables["SendKeysFromJoystick"].Select("Active=" + true + " AND Joystick='" + Joystick.joystickName[n] + "'");

            //                        for (int d = 0; d < rows.Length; d++)
            //                        {
            //                            button = rows[d]["Joystickbutton"].ToString();

            //                            if (buttonsPressed.IndexOf(button, 0) != -1) // get buttonsPressed and send keystrokes
            //                            {
            //                                sendKey = rows[d]["Keystrokes"].ToString();
            //                                optionKeyCode = rows[d]["OptionKey"].ToString();
            //                                processName = rows[d]["Processname"].ToString();

            //                                if (sendKey.Trim() != "")
            //                                {
            //                                    if (checkBoxLog.Checked)
            //                                        ImportExport.LogMessage("Send keystrokes from " + Joystick.joystickName[n] + " OptionKey: " + optionKeyCode + " Key: " + sendKey, true);

            //                                    if (optionKeyCode.Trim() != "")
            //                                        InputSimulator.SimulateModifiedKeyStroke(processName, Convert.ToUInt16(optionKeyCode), Convert.ToUInt16(sendKey));
            //                                    else
            //                                        InputSimulator.SimulateModifiedKeyStroke(processName, 0, Convert.ToUInt16(sendKey));
            //                                }
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //        catch (Exception f) { ImportExport.LogMessage("Joystick: ... " + f.ToString(), true); }

            //        lStateEnabled = true;
            //    }
            //}
            #endregion

            try
            {
                #region Switches - ADC - Encoder - Keystrokes

                if (lStateEnabled && !stop)
                {
                    lStateEnabled = false;

                    #region Switches

                    for (int n = 0; n < dataGridViewSwitches.RowCount; n++)
                    {
                        try
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
                                    if (ActivateArcaze(ref arcazeFromGrid, false))
                                    {
                                        numberInputChanged = this.arcazeHid.Command.CmdReadChangedInput();

                                        if (numberInputChanged != 0)
                                            textBoxInputChanged.Text = (numberInputChanged < 21 ? "A" + numberInputChanged.ToString() : "B" + (numberInputChanged - 20).ToString());

                                        comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesPort"];
                                        port = comboCell.Items.IndexOf(comboCell.Value);

                                        comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesPin"];
                                        pin = comboCell.Items.IndexOf(comboCell.Value);

                                        //comboCell = (DataGridViewComboBoxCell)dataGridViewSwitches.Rows[n].Cells["switchesType"];
                                        //type = comboCell.Items.IndexOf(comboCell.Value);

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
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception f) { ImportExport.LogMessage("Switches ... " + f.ToString(), true); }
                    }

                    #endregion

                    #region ADC

                    for (int n = 0; n < dataGridViewADC.RowCount; n++)
                    {
                        try
                        {
                            if (Convert.ToBoolean(dataGridViewADC.Rows[n].Cells["ADCactive"].Value))
                            {
                                if (dataGridViewADC.Rows[n].Cells["ArcazeADC"].Value != DBNull.Value &&
                                    dataGridViewADC.Rows[n].Cells["ArcazeADC"].Value != null
                                )
                                    arcazeFromGrid = dataGridViewADC.Rows[n].Cells["ArcazeADC"].Value.ToString();
                                else
                                    arcazeFromGrid = "";

                                if (arcazeFromGrid != "")
                                {
                                    if (ActivateArcaze(ref arcazeFromGrid, false))
                                    {
                                        comboCell = (DataGridViewComboBoxCell)dataGridViewADC.Rows[n].Cells["ADCChannel"];
                                        cannelADC = comboCell.Items.IndexOf(comboCell.Value);

                                        if (cannelADC > -1)
                                        {
                                            if (dataGridViewADC.Rows[n].Cells["OldValueADC"].Value != DBNull.Value)
                                                adcOldValue = Convert.ToInt16(dataGridViewADC.Rows[n].Cells["OldValueADC"].Value);
                                            else
                                                adcOldValue = 0;

                                            pinVal = ReadADC(cannelADC);

                                            if (pinVal > (adcOldValue + 2) || pinVal < (adcOldValue - 2))
                                            {
                                                dataGridViewADC.Rows[n].Cells["OldValueADC"].Value = pinVal;
                                                dataGridViewADC.Rows[n].Cells["ReadValueADC"].Value = pinVal;

                                                if (dataGridViewADC.Rows[n].Cells["TargetADC"].Value.ToString() != "")
                                                {
                                                    clickable_ID = Convert.ToInt32(dataGridViewADC.Rows[n].Cells["TargetADC"].Value);

                                                    rows = new DataRow[] { };
                                                    rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("ID=" + clickable_ID);

                                                    if (rows.Length > 0)
                                                    {
                                                        deviceID = Convert.ToInt32(rows[0]["DeviceID"]);
                                                        buttonID = Convert.ToInt32(rows[0]["ButtonID"]);
                                                        try
                                                        {
                                                            minValue_ADC = Convert.ToDouble(dataGridViewADC.Rows[n].Cells["minValueADC"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                                        }
                                                        catch
                                                        {
                                                            minValue_ADC = 0.0;
                                                            dataGridViewADC.Rows[n].Cells["minValueADC"].Value = minValue_ADC.ToString();
                                                        }
                                                        try
                                                        {
                                                            maxValue_ADC = Convert.ToDouble(dataGridViewADC.Rows[n].Cells["maxValueADC"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                                        }
                                                        catch
                                                        {
                                                            maxValue_ADC = 1.0;
                                                            dataGridViewADC.Rows[n].Cells["maxValueADC"].Value = maxValue_ADC.ToString();
                                                        }
                                                        if (minValue_ADC < -1.0)
                                                        {
                                                            minValue_ADC = -1.0;
                                                            dataGridViewADC.Rows[n].Cells["minValueADC"].Value = minValue_ADC.ToString();
                                                        }
                                                        //if (maxValue_ADC > 1.0)
                                                        //{
                                                        //    maxValue_ADC = 1.0;
                                                        //    dataGridViewADC.Rows[n].Cells["maxValueADC"].Value = maxValue_ADC.ToString();
                                                        //}
                                                        /// Ausgabewert
                                                        /// y_min = 0.2    (aus der Min Spalte)
                                                        /// y_max = 0.8    (aus der Max Spalte)
                                                        /// Eingabewert
                                                        /// x_min = 0    (fester Wert)
                                                        /// x_max = 1024    (fester Wert)

                                                        /// d_y Delta Ausgabewert (y_max - y_min) = 0.6
                                                        /// d_x Delta Eingabewert (x_max - x_min) = 1024
                                                        /// m Steigerung linearer Funktion (d_y / d_x) = 0,00058651
                                                        /// n Schnittpunkt Der Funktion mit y Achse (y_max - m * x_max) = 0,20000027

                                                        /// x Eingangswert = 512    (von der Arcaze ausgelesen)

                                                        /// y Ergebnis (m * x + n) = 0,50029339    (das wird an das Exportscript geschickt)

                                                        m_ADC = (maxValue_ADC - minValue_ADC) / 1024;
                                                        n_ADC = maxValue_ADC - (m_ADC * 1024);

                                                        dataGridViewADC.Rows[n].Cells["SendValueADC"].Value = ((Convert.ToDouble(pinVal) * m_ADC) + n_ADC).ToString();

                                                        package = "C" + deviceID + "," + (3000 + buttonID).ToString() + "," + (Convert.ToDouble(Convert.ToInt16(((Convert.ToDouble(pinVal) * m_ADC) + n_ADC) * 1000)) / 1000).ToString().Replace(",", ".");

                                                        try
                                                        {
                                                            if (checkBoxLog.Checked)
                                                                ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - ReadValue:" + pinVal.ToString("X4") + " - Package: " + package, true);

                                                            UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), package);
                                                        }
                                                        catch (Exception f)
                                                        {
                                                            ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package + " ... " + f.ToString(), true);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception f) { ImportExport.LogMessage("ADC ... " + f.ToString(), true); }
                    }
                    #endregion

                    #region Encoder

                    for (int n = 0; n < dataGridViewEncoderValues.RowCount; n++)
                    {
                        try
                        {
                            if (Convert.ToBoolean(dataGridViewEncoderValues.Rows[n].Cells["active_Encoder"].Value) &&
                                Convert.ToBoolean(dataGridViewEncoderValues.Rows[n].Cells["init_Encoders"].Value))
                            {
                                if (dataGridViewEncoderValues.Rows[n].Cells["arcaze_Encoder"].Value != DBNull.Value &&
                                    dataGridViewEncoderValues.Rows[n].Cells["arcaze_Encoder"].Value != null
                                )
                                    arcazeFromGrid = dataGridViewEncoderValues.Rows[n].Cells["arcaze_Encoder"].Value.ToString();
                                else
                                    arcazeFromGrid = "";

                                if (arcazeFromGrid != "")
                                {
                                    if (ActivateArcaze(ref arcazeFromGrid, false))
                                    {
                                        comboCell = (DataGridViewComboBoxCell)dataGridViewEncoderValues.Rows[n].Cells["encoder_Number"];
                                        encoderNo = comboCell.Items.IndexOf(comboCell.Value);

                                        if (encoderNo > -1)
                                        {
                                            encoderName = dataGridViewEncoderValues.Rows[n].Cells["encoder_Number"].Value.ToString();

                                            if (dataGridViewEncoderValues.Rows[n].Cells["oldValueEncoder"].Value != DBNull.Value)
                                                encoderOldValue = Convert.ToInt16(dataGridViewEncoderValues.Rows[n].Cells["oldValueEncoder"].Value);
                                            else
                                                encoderOldValue = 3;

                                            //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                            encoderDeltaValue = ReadEncoder(encoderNo, ref encoderOldValue);
                                            //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                                            dataGridViewEncoderValues.Rows[n].Cells["oldValueEncoder"].Value = encoderOldValue;

                                            if (encoderDeltaValue != 0)
                                            {
                                                if (dataGridViewEncoderValues.Rows[n].Cells["readValue_Encoder"].Value != DBNull.Value)
                                                    encoderLastValue = Convert.ToInt16(dataGridViewEncoderValues.Rows[n].Cells["readValue_Encoder"].Value);
                                                else
                                                    encoderLastValue = 0;

                                                try
                                                {
                                                    reverseEncoder = Convert.ToBoolean(dataGridViewEncoderValues.Rows[n].Cells["Reverse_Encoder"].Value);
                                                }
                                                catch
                                                {
                                                    reverseEncoder = false;
                                                }

                                                dataGridViewEncoderValues.Rows[n].Cells["readValue_Encoder"].Value = encoderLastValue + encoderDeltaValue;

                                                if (dataGridViewEncoderValues.Rows[n].Cells["dentValue"].Value != DBNull.Value)
                                                    encoderDentValue = double.Parse(dataGridViewEncoderValues.Rows[n].Cells["dentValue"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                                else
                                                    encoderDentValue = 1;

                                                if (dataGridViewEncoderValues.Rows[n].Cells["maxValue_Encoder"].Value != DBNull.Value)
                                                {
                                                    try
                                                    {
                                                        encoderMaxValue = double.Parse(dataGridViewEncoderValues.Rows[n].Cells["maxValue_Encoder"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                                    }
                                                    catch
                                                    {
                                                        encoderMaxValue = 10;
                                                        dataGridViewEncoderValues.Rows[n].Cells["maxValue_Encoder"].Value = encoderMaxValue;
                                                    }
                                                }
                                                else
                                                    encoderMaxValue = 10;

                                                if (dataGridViewEncoderValues.Rows[n].Cells["minValue_Encoder"].Value != DBNull.Value)
                                                {
                                                    try
                                                    {
                                                        encoderMinValue = double.Parse(dataGridViewEncoderValues.Rows[n].Cells["minValue_Encoder"].Value.ToString().Replace(",", "."), CultureInfo.InvariantCulture);
                                                    }
                                                    catch
                                                    {
                                                        encoderMinValue = 0.0;
                                                        dataGridViewEncoderValues.Rows[n].Cells["minValue_Encoder"].Value = encoderMinValue;
                                                    }
                                                }
                                                else
                                                    encoderMinValue = 0.0;

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

                                                dataGridViewEncoderValues.Rows[n].Cells["calcValue_Encoder"].Value = encoderCalcValue.ToString().Replace(",", ".");

                                                try
                                                {
                                                    encoderIdentity = Convert.ToInt16(dataGridViewEncoderValues.Rows[n].Cells["ID_Encoder"].Value);
                                                }
                                                catch
                                                {
                                                    encoderIdentity = -1;
                                                }
                                                rows = new DataRow[] { };
                                                rows = dataSetDisplaysLEDs.Tables["MultipostionSwitch"].Select("EncoderID=" + encoderIdentity.ToString());

                                                if (rows.Length > 0)
                                                {
                                                    if (rows[0]["SwitchID"] != DBNull.Value)
                                                    {
                                                        rows[0]["ValueSend"] = (reverseEncoder) ? encoderMaxValue - encoderCalcValue : encoderCalcValue;

                                                        if (Convert.ToBoolean(rows[0]["EveryValue"]))
                                                        {
                                                            rows[0]["ValueCalc"] = encoderCalcValue;

                                                            try
                                                            {
                                                                MakeDataPackageAndSend(Convert.ToInt32(rows[0]["SwitchID"]), (reverseEncoder) ? encoderMaxValue - encoderCalcValue : encoderCalcValue);
                                                            }
                                                            catch
                                                            {
                                                                ImportExport.LogMessage("Invalid Encoder - SwitchID", true);
                                                            }
                                                            dataSetDisplaysLEDs.Tables["MultipostionSwitch"].AcceptChanges();
                                                        }
                                                        else
                                                        {
                                                            if (rows[0]["ValueCalc"] != DBNull.Value)
                                                            {
                                                                if (Convert.ToInt32(rows[0]["ValueCalc"]) == encoderCalcValue)
                                                                {
                                                                    try
                                                                    {
                                                                        MakeDataPackageAndSend(Convert.ToInt32(rows[0]["SwitchID"]), double.Parse(rows[0]["ValueSend"].ToString().Replace(",", "."), CultureInfo.InvariantCulture));
                                                                    }
                                                                    catch
                                                                    {
                                                                        ImportExport.LogMessage("Invalid Encoder - SwitchID", true);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception f) { ImportExport.LogMessage("Encoder ... " + f.ToString(), true); }
                    }
                    #endregion

                    #region Keystrokes

                    for (int n = 0; n < dataGridViewKeystrokes.RowCount; n++)
                    {
                        try
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
                                    if (ActivateArcaze(ref arcazeFromGrid, false))
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
                                            try
                                            {
                                                sendKey = dataGridViewKeystrokes.Rows[n].Cells["Key"].Value.ToString();
                                                processName = dataGridViewKeystrokes.Rows[n].Cells["ProzessName"].Value.ToString();
                                                optionKeyCode = dataGridViewKeystrokes.Rows[n].Cells["OptionKeyOne"].Value.ToString();
                                                optionKeyCodeTwo = dataGridViewKeystrokes.Rows[n].Cells["OptionKeyTwo"].Value.ToString();
                                            }
                                            catch
                                            {
                                                ImportExport.LogMessage(GetActiveArcazeName() + " - ERROR OptionKeyOne: " + optionKeyCode + " OptionKeyTwo: " + optionKeyCodeTwo + " Key: " + sendKey, true);
                                                sendKey = "";
                                            }

                                            if (sendKey.Trim() != "")
                                            {
                                                if (optionKeyCode.Trim() != "" && optionKeyCodeTwo.Trim() != "")
                                                {
                                                    if (checkBoxLog.Checked)
                                                        ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  OptionKeyOne: " + optionKeyCode + " OptionKeyTwo: " + optionKeyCodeTwo + " Key: " + sendKey, true);

                                                    InputSimulator.SimulateModifiedKeyStroke(processName, UInt16.Parse(optionKeyCode.Substring(2), NumberStyles.HexNumber),
                                                        UInt16.Parse(optionKeyCodeTwo.Substring(2), NumberStyles.HexNumber), UInt16.Parse(sendKey.Substring(2), NumberStyles.HexNumber));
                                                }
                                                else
                                                {
                                                    if (optionKeyCode.Trim() != "" || optionKeyCodeTwo.Trim() != "")
                                                    {
                                                        if (optionKeyCode.Trim() != "")
                                                        {
                                                            if (checkBoxLog.Checked)
                                                                ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  OptionKeyOne: " + optionKeyCode + " Key: " + sendKey, true);

                                                            InputSimulator.SimulateModifiedKeyStroke(processName, UInt16.Parse(optionKeyCode.Substring(2), NumberStyles.HexNumber), UInt16.Parse(sendKey.Substring(2), NumberStyles.HexNumber));
                                                        }
                                                        else
                                                        {
                                                            if (checkBoxLog.Checked)
                                                                ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  OptionKeyTwo: " + optionKeyCodeTwo + " Key: " + sendKey, true);

                                                            InputSimulator.SimulateModifiedKeyStroke(processName, UInt16.Parse(optionKeyCodeTwo.Substring(2), NumberStyles.HexNumber), UInt16.Parse(sendKey.Substring(2), NumberStyles.HexNumber));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (checkBoxLog.Checked)
                                                            ImportExport.LogMessage(GetActiveArcazeName() + " - SimulateModifiedKeyStroke  Key: " + sendKey, true);

                                                        InputSimulator.SimulateModifiedKeyStroke(processName, 0, UInt16.Parse(sendKey.Substring(2), NumberStyles.HexNumber));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception f) { ImportExport.LogMessage("Keystrocke ... " + f.ToString(), true); }
                    }
                    #endregion

                    lStateEnabled = true;
                }
                #endregion

                GetLogs();

                #region Clean RAM

                loopCounter++;

                if (loopCounter == loopMax * 10)
                {
                    MemoryManagement.Reduce();
                    loopCounter = 0;
                }
                #endregion
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("Clean RAM: ... " + f.ToString(), true);
                lStateEnabled = true;
            }
        }

        #region member functions

        private string CheckLengthOfDisplayValue(ref DataRow row, string dezValue)
        {
            int displayCount = 0;
            string newValue = "";

            bool[] checkBox = new bool[8];
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

        public static string Repeat(string value, int count)
        {
            return new StringBuilder(value.Length * count).Insert(0, value, count).ToString();
        }

        private void ConvertDezHelper(ref string dezValue)
        {
            for (int n = 0; n < 8; n++)
            {
                if (checkBox[n])
                {
                    displayCount++;
                }
            }
            int foundValue = 0;

            for (int n = 0; n < dezValue.Length; n++) // Shorten the value, if to long
            {
                if (dezValue.Substring(n, 1) != ".")
                {
                    foundValue++;
                }

                if (displayCount >= foundValue)
                {
                    newValue += dezValue.Substring(n, 1);
                }
            }
            dezValue = newValue;

            for (int n = 0; n < 8; n++) // Shift the value to the left
            {
                if (!checkBox[n])
                {
                    dezValue += " ";
                }
                else
                {
                    break;
                }
            }
            int numberDigit = 0;

            for (int n = 0; n < dezValue.Length; n++) // Get number of digits
            {
                if (dezValue.Substring(n, 1) != ".")
                {
                    numberDigit++;
                }
            }
            segmentIndex = 0;

            for (int n = 0; n < dezValue.Length; n++)
            {
                if (dezValue.Substring(n, 1) == ".")
                {
                    checkBoxDP[numberDigit - segmentIndex] = true;
                }
                else
                {
                    segmentIndex++;
                }
            }
            newValue = "";

            for (int n = 0; n < dezValue.Length; n++) // Erase the dots
            {
                if (dezValue.Substring(n, 1) != ".")
                {
                    newValue += dezValue.Substring(n, 1);
                }
            }
            dezValue = newValue;

            if (useLeftPadding)
            {
                dezValue = Repeat("0", 8 - dezValue.Length) + dezValue;
            }

            for (int n = 0; n < 8; n++)
            {
                noValue = false;

                try
                {
                    if (n < checkBox.Length)
                    {
                        dezValueByte = Convert.ToInt32(dezValue.Substring((dezValue.Length - 1) - n, 1));
                    }
                    else
                    {
                        dezValueByte = 0;
                        noValue = true;
                    }
                }
                catch
                {
                    dezValueByte = 0;
                    noValue = true;
                }
                if (noValue)
                {
                    digit[n] = SevenSegment.GetNonePattern(); // <----- HJP
                }
                else
                {
                    digit[n] = ((checkBox[n] && !noValue) || (checkBox[n] && useLeftPadding && noValue)) ? SevenSegment.GetValuePattern(dezValueByte, checkBoxDP[n]) : SevenSegment.GetNonePattern();
                }
            }
        }

        private void ConvertDezToSevenSegment(ref DataRow row, ref string[] digit, string dezValue)
        {
            dezValue = dezValue.Replace("\"", "").Replace(":",".");

            if (!Convert.ToBoolean(row["Active"]))
                return;
            try
            {
                displayCount = 0;
                newValue = "";

                checkBoxDP = new bool[8] { false, false, false, false, false, false, false, false };
                checkBox = new bool[8];

                checkBox[0] = Convert.ToBoolean(row["DisplayD0"]);
                checkBox[1] = Convert.ToBoolean(row["DisplayD1"]);
                checkBox[2] = Convert.ToBoolean(row["DisplayD2"]);
                checkBox[3] = Convert.ToBoolean(row["DisplayD3"]);
                checkBox[4] = Convert.ToBoolean(row["DisplayD4"]);
                checkBox[5] = Convert.ToBoolean(row["DisplayD5"]);
                checkBox[6] = Convert.ToBoolean(row["DisplayD6"]);
                checkBox[7] = Convert.ToBoolean(row["DisplayD7"]);

                for (int n = 0; n < checkBox.Length; n++)
                {
                    try
                    {
                        if (!checkBox[n]) { digit[n] = SevenSegment.GetNonePattern(); }
                    }
                    catch { }
                }

                for (int n = 0; n < dezValue.Length; n++)
                {
                    try
                    {
                        if (dezValue[n].ToString() == " ") { digit[n] = SevenSegment.GetNonePattern(); }
                    }
                    catch { }
                }

                if (dezValue == "" || dezValue == "-")
                {
                    for (int n = 0; n < digit.Length; n++)
                    {
                        try
                        {
                            digit[n] = SevenSegment.GetNonePattern();
                        }
                        catch { }
                    }
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
                        checkBoxLog.Checked, Convert.ToBoolean(row["Reverse"]), ref segmentIndex, Convert.ToInt32(cbRefreshCycle.Text), Convert.ToInt32(cbDelayRefresh.Text));
                }
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ConvertDezToSevenSegment " + f.ToString(), true);
            }
        }

        private void ConvertDezToSevenSegment(ref string[] digit)
        {
            dezValueByte = 0;
            dezValue = textBoxDezValue.Text.Trim();
            noValue = false;
            displayCount = 0;
            newValue = "";

            checkBoxDP = new bool[8] { false, false, false, false, false, false, false, false };
            checkBox = new bool[8];

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

        private void CopyListBoxToClipboard(ref ListBox lb)
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

        private void ErasePulldown()
        {
            dataSetDisplaysLEDs.Tables["ClickableDisplay"].Clear();
            dataSetDisplaysLEDs.Tables["ClickableLED"].Clear();
            dataSetDisplaysLEDs.Tables["ClickableRotary"].Clear();
            dataSetDisplaysLEDs.Tables["ClickableSwitch"].Clear();
        }

        private void GenerateJSONDataset()
        {
            int maxRows = 62;
            string name = "";

            dtDcs_ID = dataSetDisplaysLEDs.Tables["DCS_ID"];

            try
            {
                if (configID != -1)
                {
                    dtJson = new DataTable("Data");
                    dtJson.Columns.Add("Description");
                    dtJson.Columns.Add("Type");
                    dtJson.Columns.Add("ID");
                    dtJson.Columns.Add("Format");
                    dtJson.Columns.Add("ExportID");
                    dtJson.Columns.Add("negateValue");

                    dtLamps = dataSetDisplaysLEDs.Tables["LEDs"];

                    if (dtLamps != null)
                    {
                        for (int i = 0; i < dtLamps.Rows.Count; i++)
                        {
                            try
                            {
                                DataRow[] desc = dtDcs_ID.Select("ExportID='" + dtLamps.Rows[i]["DCSExportID"].ToString() + "'");

                                name = desc[0]["Description"].ToString();

                                if (name.Length > 30) { name = name.Substring(0, 30); }

                                dataRow = dtJson.NewRow();

                                dataRow["Description"] = name;

                                dataRow["Type"] = dtLamps.Rows[i]["Type"].ToString();
                                if (dataRow["Type"].ToString() == "") { dataRow["Type"] = "ID"; }

                                dataRow["ID"] = dtLamps.Rows[i]["DCSExportID"].ToString();

                                dataRow["Format"] = dtLamps.Rows[i]["Format"].ToString();
                                if (dataRow["Format"].ToString() == "") { dataRow["Format"] = "decimal"; }

                                dataRow["ExportID"] = dtLamps.Rows[i]["DCSExportID"].ToString();

                                if (dtLamps.Rows[i]["Reverse"].ToString() == "False")
                                {
                                    dataRow["negateValue"] = "0";
                                }
                                else
                                {
                                    dataRow["negateValue"] = "1";
                                }
                                dtJson.Rows.Add(dataRow);

                                if (dtJson.Rows.Count > maxRows)
                                {
                                    dtJson.AcceptChanges();

                                    GenerateAndSentJson(dtJson);

                                    dtJson.Clear();

                                    Thread.Sleep(10);
                                }
                            }
                            catch (Exception f)
                            {
                                ImportExport.LogMessage("GenerateJSONDataset for lamps " + f.ToString(), true);
                            }
                        }
                    }

                    dtDisplays = dataSetDisplaysLEDs.Tables["Displays"];

                    if (dtDisplays != null)
                    {
                        for (int i = 0; i < dtDisplays.Rows.Count; i++)
                        {
                            try
                            {
                                if (dtDisplays.Rows[i]["DCSExportID"].ToString() != "")
                                {
                                    DataRow[] desc = dtDcs_ID.Select("ExportID='" + dtDisplays.Rows[i]["DCSExportID"].ToString() + "'");

                                    name = desc[0]["Description"].ToString();

                                    if (name.Length > 30) { name = name.Substring(0, 30); }

                                    dataRow = dtJson.NewRow();
                                    dataRow["Description"] = name;

                                    dataRow["Type"] = dtDisplays.Rows[i]["Type"].ToString();
                                    if (dataRow["Type"].ToString() == "") { dataRow["Type"] = "ID"; }

                                    dataRow["ID"] = dtDisplays.Rows[i]["IDExportscript"].ToString();
                                    if (dataRow["ID"].ToString() == "") { dataRow["ID"] = dtDisplays.Rows[i]["DCSExportID"].ToString(); }

                                    dataRow["Format"] = dtDisplays.Rows[i]["Format"].ToString();
                                    dataRow["ExportID"] = dtDisplays.Rows[i]["DCSExportID"].ToString();

                                    if (dtDisplays.Rows[i]["Reverse"].ToString() == "False")
                                    {
                                        dataRow["negateValue"] = "0";
                                    }
                                    else
                                    {
                                        dataRow["negateValue"] = "1";
                                    }

                                    dtJson.Rows.Add(dataRow);

                                    if (dtJson.Rows.Count > maxRows)
                                    {
                                        dtJson.AcceptChanges();

                                        GenerateAndSentJson(dtJson);

                                        dtJson.Clear();

                                        Thread.Sleep(10);
                                    }
                                }
                            }
                            catch (Exception f)
                            {
                                ImportExport.LogMessage("GenerateJSONDataset for displays: " + f.ToString(), true);
                            }
                        }
                    }

                    if (dtJson.Rows.Count > 0)
                    {
                        try
                        {
                            dtJson.AcceptChanges();

                            GenerateAndSentJson(dtJson);

                            dtJson.Clear();
                        }
                        catch (Exception f)
                        {
                            ImportExport.LogMessage("GenerateJSONDataset for switches: " + f.ToString(), true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ImportExport.LogMessage("GenerateJSONDataset: " + ex.ToString(), true);
            }
        }

        private void GenerateAndSentJson(DataTable dtJson)
        {
            try
            {
                if (dtJson.Rows.Count > 0)
                {
                    dsJSON = new DataSet();
                    dsJSON.Tables.Add(dtJson);

                    json = JsonConvert.SerializeObject(dsJSON, Formatting.None);

                    dsJSON.Tables.Remove(dtJson);

                    string configIDString = "{'ConfigID': " + configID + ", ";
                    configIDString = configIDString.Replace("'", '"'.ToString());

                    json = configIDString + json.Substring(1, json.Length - 1);

                    UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), json);

                    ImportExport.LogMessage("Send json data -> " + json.ToString() + " - " + json.Length + " bytes.", true);
                }
            }
            catch (Exception ex)
            {
                ImportExport.LogMessage("GenerateAndSentJson: " + ex.ToString(), true);
            }
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

                    if (listBox1.Items.Count > 2000)
                        listBox1.Items.RemoveAt(6);
                }
                try
                {
                    ImportExport.log.Clear();
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

        private void CockpitLoad(ref string receivedData)
        {
            try
            {
                if (GetIdent(searchStringForFile, ref receivedData))
                {
                    if (readFile.Length > 0)
                    {
                        readFile += ".xml";
                        loadFile = true;
                        timerMain.Interval = timerInterval / 5;
                        timerstate = State.readNewFile;

                        try
                        {
                            receivedDataBackup = receivedData.Substring(receivedData.IndexOf(searchStringForFile) + 5);
                            receivedData = receivedDataBackup;
                        }
                        catch
                        {
                            receivedDataBackup = "";
                        }
                    }
                }
            }
            catch (Exception ex) { ImportExport.LogMessage("Load Cockpit: " + ex.ToString(), true); }
        }

        private bool GetIdent(string ID, ref string gotData)
        {
            string[] receivedItems = gotData.Split(':');

            try
            {
                for (int n = 0; n < receivedItems.Length; n++)
                {
                    if (receivedItems[n].IndexOf(ID, 0) == 0)
                    {
                        readFile = receivedItems[n].Substring(receivedItems[n].IndexOf("=", 0) + 1);
                        readFile = readFile.Replace('"'.ToString(), "");
                        ImportExport.LogMessage("DCS start command for modul: " + readFile, true);
                        json = "";
                        return true;
                    }
                }
            }
            catch (Exception e) { ImportExport.LogMessage("GetCockpit problem .. " + e.ToString(), true); }

            return false;
        }

        private void GetConfigID(ref string gotData)
        {
            string[] receivedItems = gotData.Split(':');

            try
            {
                for (int n = 0; n < receivedItems.Length; n++)
                {
                    if (receivedItems[n].IndexOf("ConfigID=") != -1)
                    {
                        configID = Convert.ToInt32(receivedItems[n].Substring(receivedItems[n].IndexOf("=", 0) + 1));
                        break;
                    }
                }
            }
            catch (Exception e) { ImportExport.LogMessage("GetConfigID problem .. " + e.ToString(), true); }
        }

        public void CheckReceivedItem()
        {
            for (int i = 0; i < receivedItems.Length; i++)
            {
                try
                {
                    if (receivedItems[i].IndexOf(start) > 0)
                    {
                        for (int n = 1; n < receivedItems.Length - i; n++)
                        {
                            if (receivedItems[i + n].IndexOf(start) > 0 || receivedItems[i + n].IndexOf("=0") > 0 || receivedItems[i + n].IndexOf("=1") > 0 ||
                                     receivedItems[i + n].IndexOf("=-0") > 0 || receivedItems[i + n].IndexOf("=-1") > 0 || receivedItems[i + n].IndexOf("=-2") > 0 ||
                                     receivedItems[i + n].IndexOf("=2") > 0 || receivedItems[i + n].IndexOf("=3") > 0 || receivedItems[i + n].IndexOf("=4") > 0 ||
                                     receivedItems[i + n].IndexOf("=5") > 0 || receivedItems[i + n].IndexOf("=6") > 0 || receivedItems[i + n].IndexOf("=7") > 0 ||
                                     receivedItems[i + n].IndexOf("=8") > 0 || receivedItems[i + n].IndexOf("=9") > 0 || receivedItems[i + n].IndexOf("=\"") > 0)
                            {
                                break;
                            }
                            else
                            {
                                if (receivedItems[i + n] != "" && receivedItems[i + n] != "\r\n")
                                {
                                    receivedItems[i] += ":" + receivedItems[i + n];
                                    receivedItems[i + n] = "";
                                }
                            }
                        }
                    }
                }
                catch
                { }
            }
        }

        private string GrabValue()
        {
            try
            {
                idFound = false;

                if (receivedData.IndexOf(dcsExportID, 0) > -1) // Dirty quickcheck before loop
                {
                    idFound = true;

                    for (int n = 0; n < receivedItems.Length; n++)
                    {
                        if (receivedItems[n].IndexOf(dcsExportID, 0) == 0)
                        {
                            return receivedItems[n].Substring(receivedItems[n].IndexOf("=", 0) + 1).Replace(("\r"), string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                ImportExport.LogMessage("GrabValue problem .. " + e.ToString(), true);
            }

           return "";
        }

        private void GrabValues()
        {
            arcazeFromGrid = "";

            if (receivedData.IndexOf("ConfigID=") != -1)
            {
                GetConfigID(ref receivedData);
            }

            if (receivedData.IndexOf(searchStringForFile) != -1)
            {
                CockpitLoad(ref receivedData);
            }

            if (loadFile)
            {
                return;
            }
            if (receivedData.IndexOf(@"DCS=""Stop""") != -1)
            {
                ImportExport.LogMessage(@"DCS=""Stop""", true);

                ImportExport.LogMessage("Reset configID = -1", true);

                configID = -1;
                timerstate = State.init;

                MemoryManagement.Reduce();
                return;
            }

            receivedItems = receivedData.Split(':');

            CheckReceivedItem();

            if (dataGridViewArcaze.Rows.Count > 0)
            {
                #region Displays

                for (int n = 0; n < dataGridViewDisplays.RowCount; n++)
                {
                    if (Convert.ToBoolean(dataGridViewDisplays.Rows[n].Cells["activeDisplaysCheckBox"].Value) &&
                        Convert.ToBoolean(dataGridViewDisplays.Rows[n].Cells["displayInit"].Value))
                    {
                        if (dataGridViewDisplays.Rows[n].Cells["arcazeDisplays"].Value != DBNull.Value &&
                            dataGridViewDisplays.Rows[n].Cells["arcazeDisplays"].Value != null
                        )
                        {
                            arcazeFromGrid = dataGridViewDisplays.Rows[n].Cells["arcazeDisplays"].Value.ToString();
                        }
                        else
                        {
                            arcazeFromGrid = "";
                        }
                        if (arcazeFromGrid != "")
                        {
                            if (dataGridViewDisplays.Rows[n].Cells["dcsExportIDDisplays"].Value != DBNull.Value)
                            {
                                dcsExportID = dataGridViewDisplays.Rows[n].Cells["dcsExportIDDisplays"].Value.ToString();
                            }
                            else
                            {
                                dcsExportID = "";
                            }
                            if (dcsExportID != "")
                            {
                                newValue = GrabValue();
                                newValue = newValue.Replace("\"", "");
                            }
                            else
                            {
                                newValue = "";
                            }

                            if (newValue != "")
                            {
                                if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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
                            {
                                newValue = GrabValue();
                                newValue = newValue.Replace("\"", "");
                            }
                            else
                                newValue = "";

                            if (newValue != "")
                            {
                                if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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

                                        if (dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value != DBNull.Value)
                                        {
                                            pinMax = dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value.ToString().Replace(",", ".");
                                        }
                                        else
                                        {
                                            pinMax = "1.0";
                                            dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value = "1.0";
                                        }

                                        try
                                        {
                                            pinValue = double.Parse(pinString, CultureInfo.InvariantCulture);
                                            arcazeDeviceIndex = Convert.ToInt32(dataGridViewLEDs.Rows[n].Cells["deviceIndex"].Value);

                                            switch (type)
                                            {
                                                case 2: // LED Driver 2
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port, ref pin, ref resolution, Convert.ToDouble(pinValue > 0.5 ? 1 : 0), type, isReverse, checkBoxLog.Checked);
                                                    break;

                                                case 3: // LED Driver 3
                                                    pinValueMax = double.Parse(pinMax, CultureInfo.InvariantCulture);

                                                    if (pinValueMax > 1.0) { pinValueMax = 1.0; }
                                                    if (pinValueMax < 0.0) { pinValueMax = 0.0; }

                                                    pinValue *= pinValueMax;
                                                    dataGridViewLEDs.Rows[n].Cells["LED_Max"].Value = pinValueMax.ToString();

                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port, ref pin, ref resolution, pinValue, type, isReverse, checkBoxLog.Checked);
                                                    break;

                                                case 4: // Arcase USB
                                                    arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port, ref pin, ref resolution, Convert.ToDouble(pinValue > 0.5 ? 1 : 0), type, isReverse, checkBoxLog.Checked);
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
            if (!debug)
            {
                if (arcazeDevice != null)
                    arcazeDevice.Clear();
            }
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
                    if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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
            try
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

                dataSetConfig.Tables["VirtualKeys"].Rows.Add(" ", "0x00");
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
            catch { }
        }

        private void InitEncoders()
        {
            for (int n = 0; n < dataGridViewEncoderValues.RowCount; n++)
            {
                dataGridViewEncoderValues.Rows[n].Cells["init_Encoders"].Value = false;
                dataGridViewEncoderValues.Rows[n].Cells["ReadValue_Encoder"].Value = 0;
                dataGridViewEncoderValues.Rows[n].Cells["CalcValue_Encoder"].Value = 0;

                if (dataGridViewEncoderValues.Rows[n].Cells["Arcaze_Encoder"].Value != DBNull.Value &&
                    dataGridViewEncoderValues.Rows[n].Cells["Arcaze_Encoder"].Value != null
                )
                    arcazeFromGrid = dataGridViewEncoderValues.Rows[n].Cells["Arcaze_Encoder"].Value.ToString();
                else
                    arcazeFromGrid = "";

                if (arcazeFromGrid != "")
                {
                    if (ActivateArcaze(ref arcazeFromGrid, false))
                    {
                        comboCell = (DataGridViewComboBoxCell)dataGridViewEncoderValues.Rows[n].Cells["encoder_Number"];
                        encoderNo = comboCell.Items.IndexOf(comboCell.Value);

                        if (encoderNo > -1)
                        {
                            GetPortPinOfEncoder(ref encoderNo, ref port, ref pin);

                            ImportExport.LogMessage("Init Encoder " + (encoderNo + 1).ToString("D2"), true);

                            SetPinDirection(port, pin, 0);          // direction input
                            SetPinDirection(port, pin + 1, 0);

                            dataGridViewEncoderValues.Rows[n].Cells["init_Encoders"].Value = true;
                            dataGridViewEncoderValues.Rows[n].Cells["oldValueEncoder"].Value = 3;
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
                    if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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
                    if (ActivateArcaze(ref arcaze, false))
                    {
                        for (int m = 0; m < arcazeDevice.Count; m++)
                        {
                            if (arcazeDevice[m].GetSerial == arcazeHid.Info.Serial)
                            {
                                return m;
                            }
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
                    if (ActivateArcaze(ref arcaze, false))
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
                    if (ActivateArcaze(ref arcazeFromGrid, false))
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
                        if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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
                    if (ActivateArcaze(ref arcazeFromGrid, false)) // Activate the arcaze, if necessary
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
                    {
                        ImportExport.LogMessage(GetActiveArcazeName() + " - Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package, true);
                    }
                    UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), package);
                }
                catch (Exception f)
                {
                    ImportExport.LogMessage("Send package to IP: " + textBoxIP.Text.Trim() + " - Port: " + textBoxPortSender.Text + " - Package: " + package + " ... " + f.ToString(), true);
                }
            }
        }

        private void RefreshAllDatagrids()
        {
            try
            {
                dataGridViewDisplays.Refresh();
                dataGridViewDisplays.Update();

                dataGridViewLEDs.Refresh();
                dataGridViewLEDs.Update();

                dataGridViewSwitches.Refresh();
                dataGridViewSwitches.Update();

                dataGridViewEncoderValues.Refresh();
                dataGridViewEncoderValues.Update();

                dataGridViewEncoderSend.Refresh();
                dataGridViewEncoderSend.Update();

                dataGridViewADC.Refresh();
                dataGridViewADC.Update();
            }
            catch { }
        }

        private void RefreshPulldown()
        {
            RefreshExportDisplay();
            RefreshExportLED();
            RefeshClickableRotary();
            RefeshClickableSwitch();

            RefreshAllDatagrids();
        }

        private void RefreshExportDisplay()
        {
            dataSetDisplaysLEDs.Tables["ClickableDisplay"].Clear();

            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["DCS_ID"].Select("Type='Display'");

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables["ClickableDisplay"].Rows.Add(rows[n][0], rows[n][2]);
            }
            dataSetDisplaysLEDs.Tables["ClickableDisplay"].AcceptChanges();
        }

        private void RefreshExportLED()
        {
            dataSetDisplaysLEDs.Tables["ClickableLED"].Clear();

            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["DCS_ID"].Select("Type='Lamp'");

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables["ClickableLED"].Rows.Add(rows[n][0], rows[n][2]);
            }
            dataSetDisplaysLEDs.Tables["ClickableLED"].AcceptChanges();
        }

        private void RefeshClickableRotary()
        {
            dataSetDisplaysLEDs.Tables["ClickableRotary"].Clear();

            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("Type='Rotary'");

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables["ClickableRotary"].Rows.Add(rows[n][0], rows[n][3]);
            }
            dataSetDisplaysLEDs.Tables["ClickableRotary"].AcceptChanges();
        }

        private void RefeshClickableSwitch()
        {
            dataSetDisplaysLEDs.Tables["ClickableSwitch"].Clear();

            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("Type='Switch'");

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables["ClickableSwitch"].Rows.Add(rows[n][0], rows[n][3]);
            }
            dataSetDisplaysLEDs.Tables["ClickableSwitch"].AcceptChanges();
        }

        private void SaveConfig()
        {
            if (dataSetConfig.Tables["Config"].Rows.Count > 0)
            {
                dataSetConfig.Tables["Config"].Rows[0]["LastFile"] = textBoxLastFile.Text;
                dataSetConfig.Tables["Config"].Rows[0]["WriteToHD"] = false;
                dataSetConfig.Tables["Config"].Rows[0]["LogAllActions"] = false;
                dataSetConfig.Tables["Config"].Rows[0]["PortListener"] = textBoxPortListener.Text;
                dataSetConfig.Tables["Config"].Rows[0]["PortSender"] = textBoxPortSender.Text;
                dataSetConfig.Tables["Config"].Rows[0]["IntervalTimer"] = textBoxIntervalTimer.Text;
                dataSetConfig.Tables["Config"].Rows[0]["Tlc5943Brightness"] = trackBarLEDDriverBrightness.Value.ToString("X2");
                dataSetConfig.Tables["Config"].Rows[0]["Display7219DigitsIntensity"] = trackBarDisplayBrightness.Value.ToString("X2");
                dataSetConfig.Tables["Config"].Rows[0]["TestData"] = textBoxTestDataPackage.Text.Replace(newline, "");
                dataSetConfig.Tables["Config"].Rows[0]["refreshDelay"] = cbDelayRefresh.Text;
                dataSetConfig.Tables["Config"].Rows[0]["refreshCycles"] = cbRefreshCycle.Text;
            }
            else
            {
                dataSetConfig.Tables["Config"].Rows.Add(textBoxLastFile.Text, textBoxIP.Text, textBoxPortListener.Text, trackBarLEDDriverBrightness.Value.ToString("X2"), trackBarDisplayBrightness.Value.ToString("X2"),
                    textBoxTestDataPackage.Text.Replace(newline, ""), textBoxIntervalTimer.Text, false, false, textBoxPortSender.Text, cbRefreshCycle.Text, cbDelayRefresh.Text);
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
            UDP.StartListener(Convert.ToInt16(textBoxPortListener.Text.Trim()));
        }

        private void StartTimer()
        {
            timerInterval = Convert.ToInt32(textBoxIntervalTimer.Text);
            timerMain.Interval = timerInterval;
            timerMain.Tick += new EventHandler(TimerMain_Tick);
            timerMain.Start();
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
                        receivedData = ":" + dcsExportID + "\"= 1\":";

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
                        receivedData = ":" + dcsExportID + "=\"" + (onValue ? "1" : "0") + "\":";

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
                        receivedData = ":" + dcsExportID + "=\"" + (onValue ? "88888888" : "-") + "\":";
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
        }

        private void SystemCheck()
        {
            buttonSendTestPattern.Enabled = false;
            testLoop = 0;
            ledOn = !ledOn;
            labelOnOff.Text = (ledOn ? "Switch 'Off'" : "Switch 'On'");
            timerstate = State.checkLEDS;
        }

        private void FillTypeClickable()
        {
            bool refreshDatabase = false;
            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select();

            for (int n = 0; n < rows.Length; n++)
            {
                if (rows[n]["Type"] == DBNull.Value)
                {
                    rows[n]["Type"] = "Switch";
                    refreshDatabase = true;
                }
            }

            if (refreshDatabase)
            {
                //---------------
                // Encoder
                DataRow[] rows2 = new DataRow[] { };
                rows2 = dataSetDisplaysLEDs.Tables["MultipostionSwitch"].Select();

                for (int n = 0; n < rows2.Length; n++)
                {
                    rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("ID='" + rows2[n][1] + "'");

                    if (rows.Length > 0)
                    {
                        if (rows[0]["Type"].ToString() != "Rotary")
                        {
                            rows[0]["Type"] = "Rotary";
                        }
                    }
                }
                //--------------
                // ADC
                rows2 = new DataRow[] { };
                rows2 = dataSetDisplaysLEDs.Tables["ADC"].Select();

                for (int n = 0; n < rows2.Length; n++)
                {
                    rows = dataSetDisplaysLEDs.Tables["Clickabledata"].Select("ID='" + rows2[n][10] + "'");

                    if (rows.Length > 0)
                    {
                        if (rows[0]["Type"].ToString() != "Rotary")
                        {
                            rows[0]["Type"] = "Rotary";
                        }
                    }
                }
            }
            dataSetDisplaysLEDs.Tables["Clickabledata"].AcceptChanges();
        }

        private void FillTypeDCSID()
        {
            bool refreshDatabase = false;
            rows = new DataRow[] { };
            rows = dataSetDisplaysLEDs.Tables["DCS_ID"].Select();

            for (int n = 0; n < rows.Length; n++)
            {
                if (rows[n]["Type"] == DBNull.Value)
                {
                    rows[n]["Type"] = "Lamp";
                    refreshDatabase = true;
                }
            }
            if (refreshDatabase)
            {
                //--------------
                // Display
                DataRow[] rows2 = new DataRow[] { };
                rows2 = dataSetDisplaysLEDs.Tables["Displays"].Select();

                for (int n = 0; n < rows2.Length; n++)
                {
                    temp = rows2[n][1].ToString();
                    rows = dataSetDisplaysLEDs.Tables["DCS_ID"].Select("ExportID='" + rows2[n][1] + "'");

                    if (rows.Length > 0)
                    {
                        if (rows[0]["Type"].ToString() != "Display") { rows[0]["Type"] = "Display"; }
                    }
                }
            }
            dataSetDisplaysLEDs.Tables["DCS_ID"].AcceptChanges();
        }

        #endregion

        #region Arcaze

        private bool ActivateArcaze(ref string arcazeNew, bool log = true)
        {
            if (arcazeNew == "")
            {
                return false;
            }

            if (arcazeAddress == arcazeNew)
            {
                return true;
            }

            rows = new DataRow[] { };
            //rows = dataSetDisplaysLEDs.Tables["Arcaze"].Select("Arcaze='" + arcazeNew + "'");
            rows = dataSetDisplaysLEDs.Tables["Arcaze"].Select("SerialNumber='" + arcazeNew + "'");

            if (rows.Length == 0)
            {
                return false;
            }

            for (int n = 0; n < rows.Length; n++)
            {
                arcazeActive = Convert.ToBoolean(rows[n]["Active"]);

                if (!arcazeActive)
                {
                    return false;
                }
            }

            if (arcazeNew != arcazeAddress)
            {
                for (int n = 0; n < presentArcaze.Count; n++)
                {
                    if (arcazeNew == presentArcaze[n].Serial)
                    {
                        try
                        {
                            if (checkBoxLog.Checked && arcazeHid.Info.Connected && log)
                            {
                                ImportExport.LogMessage("Disconnect " + arcazeHid.Info.DeviceName + " (" + arcazeHid.Info.Serial + ")", true);
                            }

                            if (arcazeHid.Info.Connected)
                            {
                                arcazeHid.Disconnect();
                            }

                            if (checkBoxLog.Checked && log)
                            {
                                ImportExport.LogMessage("Connect " + presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")", true);
                            }

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
            {
                return true;
            }
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
            {
                arcazeHid.Disconnect();
            }
        }

        private ArcazeCommand.OutputOperators GetOutputOperator()
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
            allArcaze = new List<DeviceInfo>(16);

            if (debug)
            {
                presentArcaze.Clear();
                presentArcaze.Add(new DeviceInfo());
                presentArcaze.Add(new DeviceInfo());
                presentArcaze[0].Serial = "000413300000";
                presentArcaze[0].DeviceName = "000413300000";
                presentArcaze[1].Serial = "000414900000";
                presentArcaze[1].DeviceName = "000414900000";
            }
            else
            {
                arcazeHid.Find(allArcaze);
                presentArcaze = arcazeHid.RemoveSameSerialDevices(allArcaze);
            }
            ComboBoxArcaze.Items.Clear();
            ComboBoxArcaze.Text = "";

            findArcaze = new DataRow[] { };
            dataSetDisplaysLEDs.Tables["Arcaze"].Clear();

            ComboBoxArcaze.Text = "";

            for (int n = 0; n < presentArcaze.Count; n++)
            {
                //ComboBoxArcaze.Items.Add(presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")");
                ComboBoxArcaze.Items.Add(presentArcaze[n].Serial);

                row = dataSetDisplaysLEDs.Tables["Arcaze"].NewRow();
                row[0] = presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")";
                row[1] = true;
                row[2] = presentArcaze[n].Serial;

                try
                {
                    dataSetDisplaysLEDs.Tables["Arcaze"].Rows.Add(row);
                }
                catch { }

                RepairDatabaseForArcaze(presentArcaze[n].DeviceName + " (" + presentArcaze[n].Serial + ")", presentArcaze[n].Serial);
            }
            dataSetDisplaysLEDs.Tables["Arcaze"].AcceptChanges();


            if (presentArcaze.Count > 0)
            {
                ImportExport.LogMessage(presentArcaze.Count.ToString() + " Arcaze found: ", true);

                for (int n = 0; n < presentArcaze.Count; n++)
                {
                    ImportExport.LogMessage("\t\t Serialnumber: " + presentArcaze[n].Serial + " - Interface: " + presentArcaze[n].Interface, false);
                }

                ComboBoxArcaze.SelectedIndex = 0;
                arcazeHid.Connect(presentArcaze[0].Path);
            }
            return presentArcaze.Count > 0;
        }

        private void RepairDatabaseForArcaze(string arcazeName, string serialNumber)
        {
            rows = dataSetDisplaysLEDs.Tables[0].Select("Arcaze='" + arcazeName + "'"); // LEDs

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables[0].Rows[n][6] = serialNumber; // Field Arcase
            }

            rows = dataSetDisplaysLEDs.Tables[2].Select("Arcaze='" + arcazeName + "'"); // Displays

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables[2].Rows[n][11] = serialNumber; // Field Arcase
            }

            rows = dataSetDisplaysLEDs.Tables[6].Select("Arcaze='" + arcazeName + "'"); // Switches

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables[6].Rows[n][1] = serialNumber; // Field Arcase
            }

            rows = dataSetDisplaysLEDs.Tables[9].Select("Arcaze='" + arcazeName + "'"); // Encoders

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables[9].Rows[n][3] = serialNumber; // Field Arcase
            }

            rows = dataSetDisplaysLEDs.Tables[11].Select("Arcaze='" + arcazeName + "'"); // Keystrokes

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables[11].Rows[n][2] = serialNumber; // Field Arcase
            }

            rows = dataSetDisplaysLEDs.Tables[12].Select("Arcaze='" + arcazeName + "'"); // ADC

            for (int n = 0; n < rows.Length; n++)
            {
                dataSetDisplaysLEDs.Tables[12].Rows[n][3] = serialNumber; // Field Arcase
            }

            dataSetDisplaysLEDs.AcceptChanges();
        }

        private string GetActiveArcazeName()
        {
            //return arcazeHid.Info.DeviceName + " (" + arcazeHid.Info.Serial + ")";
            return arcazeHid.Info.Serial;
        }

        private void ResetArcaze()
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                return;

            try
            {
                if (checkBoxLog.Checked)
                {
                    ImportExport.LogMessage("CmdReset() ", true);
                }

                arcazeHid.Command.CmdReset();
            }
            catch (Exception e)
            {
                ImportExport.LogMessage("arcazeHid.ArcazeCommands.Reset() .. " + e.ToString(), true);
            }
        }

        #endregion

        #region LED / ADC / Encoder

        //private int GetConnectorNumber(int modulNumber, int connector)
        //{
        //    if (modulNumber == 0)
        //        return (connector == 0 ? 0 : 1);
        //    else
        //        return ((connector - 2) + (3 * (modulNumber - 1)));
        //}

        //private void LEDWriteOutputOld(int moduleNum, int connectorNum, int portNum, int resolution, Double data, int type, bool reverse)
        //{
        //    if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
        //        return;

        //    if (type != 3)
        //    {
        //        if (data < 0.485 && data > 0.001)
        //            return;
        //        if (data > 0.515 && data < 0.999)
        //            return;

        //        data = (data > 0.5 ? 1 : 0); // Arcaze and LED-Driver 2

        //        if (reverse)
        //            data = (data == 1 ? 0 : 1);
        //    }
        //    else
        //    {
        //        if (reverse)
        //            data = 1 - data;

        //        resolutionValue = Convert.ToUInt32(ResolutionValue(resolution)); // LED-Driver 3
        //        data *= resolutionValue;

        //        if (data > resolutionValue)
        //            data = resolutionValue;

        //        if (data < 0)
        //            data = 0;
        //    }

        //    if (connectorNum > 1)
        //        connectorNum -= 2;

        //    //connectorNum = GetConnectorNumber(moduleNum, connectorNum);

        //    try
        //    {
        //        if (checkBoxLog.Checked)
        //            ImportExport.LogMessage("WriteOutputPort(Modul: " + moduleNum.ToString("X2") + ", Connector: " + connectorNum.ToString("X2") + ", Pin: " + (portNum + 1).ToString("D2") + ", Value: " + (data == 0 ? "Off" : (type != 3 ? "On" : data.ToString())) + ")", true);

        //        this.arcazeHid.Command.WriteOutputPort(moduleNum, connectorNum, portNum, Convert.ToUInt32(data), ArcazeCommand.OutputOperators.PlainWrite, false);
        //        this.arcazeHid.Command.UpdateOutputPorts();
        //    }
        //    catch (Exception e)
        //    {
        //        ImportExport.LogMessage("WriteOutputPort(Modul: " + moduleNum.ToString("X2") + ", Connector: " + connectorNum.ToString("X2") + ", Pin: " + (portNum + 1).ToString("D2") + ", Value: " + (data == 0 ? "Off" : (moduleNum == 0 ? "On" : data.ToString())) + ") ... " + e.ToString(), true);
        //    }
        //}

        private int ReadADC(int channel)
        {

            int[] inputs = new int[6];

            try
            {
                if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                    return 0;

                //arcazeHid.Command.CmdPrintADC();
                inputs = arcazeHid.Command.CmdReadADC();

                //if (checkBoxLog.Checked)
                //    ImportExport.LogMessage("ReadADC(" + channel.ToString("D2") + "), Value = " + inputs[channel].ToString("X4"), true);

                return inputs[channel];
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ReadADC(" + channel.ToString("D2") + ") ... " + f.ToString(), true);
                return 0;
            }
        }

        private int[] ReadADC()
        {
            int[] inputs = new int[6];

            try
            {
                if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
                    return inputs;

                //arcazeHid.Command.CmdPrintADC();
                inputs = arcazeHid.Command.CmdReadADC();

                for (int i = 0; i < 6; i++)
                {
                    if (checkBoxLog.Checked)
                        ImportExport.LogMessage("ReadADC(" + i.ToString("D2") + "), Value = " + inputs[i].ToString("X4"), true);
                }
                return inputs;
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ReadADC ... " + f.ToString(), true);
                return inputs;
            }
        }

        private int ReadEncoder(int encoderNumber, ref int encoderOldValue)
        {
            short[] encoderValues = new short[20];
            //arcazeHid.Command.CmdReadQuadratureEncodersRelative(ref encoderValues);
            arcazeHid.Command.CmdReadQuadratureEncodersAbsolute(ref encoderValues);
            encoderNewValue = encoderValues[encoderNumber];

            encoderReturnValue = (encoderNewValue - encoderOldValue) / 2;
            encoderOldValue = encoderNewValue; // / 4;
            return encoderReturnValue;
        }

        private int ReadEncoderOldway(int encoderNumber, ref int encoderOldValue)
        {
            GetPortPinOfEncoder(ref encoderNumber, ref port, ref pin);


            ReadEncoderPinsValues(port, pin, ref encoderPinAvalue, ref encoderPinBvalue);

            encoderNewValue = (encoderPinAvalue * 2) + encoderPinBvalue;

            if (encoderNewValue == encoderOldValue)
            {
                return 0;
            }

            encoderOldValue = 3;

            encoderReturnValue = RotaryEncoder.Encoder(ref encoderNewValue, ref encoderOldValue);
            encoderOldValue = encoderNewValue;

            return encoderReturnValue;
        }

        private void ReadEncoderPinsValues(int port, int pin, ref int pinAvalue, ref int pinBvalue)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
            {
                return;
            }

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
                {
                    return;
                }

                value = arcazeHid.Command.CmdReadPort(port);

                value = value >> pin;
                pattern = Convert.ToString(value, 2);
                value = Convert.ToInt32(pattern.Substring(pattern.Length - 1));

                if (!reverse)
                {
                    value = (value == 1 ? 0 : 1); // Invert it
                }
            }
            catch (Exception f)
            {
                ImportExport.LogMessage("ReadPort(" + port + ") .. " + f.ToString(), true);
            }
        }

        private void SetPinDirection(int port, int pin, int direction)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected || !arcazeFound)
            {
                return;
            }

            try
            {
                ImportExport.LogMessage(arcazeHid.Info.Serial + " SetPinDirection(Connector: " + port.ToString("X2") + ", Pin: " + (pin + 1).ToString("D2") + ", Direction: " + (direction == 1 ? "Output" : "Input") + ")", true);

                arcazeHid.Command.CmdSetPinDirection(port, pin, direction);
            }
            catch (Exception f)
            {
                ImportExport.LogMessage(arcazeHid.Info.Serial + " SetPinDirection(Connector: " + port.ToString("X2") + ", Pin: " + (pin + 1).ToString("D2") + ", Direction: " + (direction == 1 ? "Output" : "Input") + ") .. " + f.ToString(), true);
            }
        }

        #endregion

        #region events

        private void ArcazeChanged(object sender, EventArgs e)
        {
            ChangeArcase(presentArcaze[ComboBoxArcaze.SelectedIndex].Path);
        }

        private void ButtonCopyToClipboard_Click(object sender, EventArgs e)
        {
            CopyListBoxToClipboard(ref listBox1);
        }

        private void ButtonDonate_Click(object sender, EventArgs e)
        {
            string url = "";

            string business = "heinz.joerg.puhlmann@googlemail.com";  // your paypal email
            string description = "Donation";            // '%20' represents a space. remember HTML!
            string country = "DE";                  // AU, US, etc.
            string currency = "EUR";                 // AUD, USD, etc.

            url += "https://www.paypal.com/cgi-bin/webscr" +
                "?cmd=" + "_donations" +
                "&business=" + business +
                "&lc=" + country +
                "&item_name=" + description +
                "&currency_code=" + currency +
                "&bn=" + "PP%2dDonationsBF";

            System.Diagnostics.Process.Start(url);
        }

        private void ButtonFind_Click(object sender, EventArgs e)
        {
            InitAll();
        }

        private void ButtonFillType_Click(object sender, EventArgs e)
        {
            FillTypeClickable();
        }

        private void ButtonInit_Click(object sender, EventArgs e)
        {
            buttonInit.Visible = false;
            InitAll();
        }

        private void ButtonLogClear_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            ImportExport.LogMessage(labelVersion.Text, true);
            ImportExport.LogMessage("Used configuration file '" + textBoxLastFile.Text + "'", true);
        }

        private void ButtonOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Application.StartupPath,
                Filter = "D.A.C. Data Files *.xml|*.xml|*.*|*.*",
                Title = "Select a .xml File"
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SwitchInverseOff();

                textBoxLastFile.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\', openFileDialog1.FileName.Length - 1) + 1);
                ImportExport.LogMessage("Used configuration file " + textBoxLastFile.Text, true);
                lastFile = textBoxLastFile.Text.Substring(0, textBoxLastFile.Text.LastIndexOf("."));
                readFile = lastFile;

                dataSetFilename = openFileDialog1.FileName;

                dataSetDisplaysLEDs.Clear();
                ImportExport.XmlToDataSet(openFileDialog1.FileName, dataSetDisplaysLEDs);
                ImportExport.DatasetToXml("config.xml", dataSetConfig);

                RefreshPulldown();
                InitTables();
                InitAll();

                this.Text = this.Text.Substring(0, (this.Text.IndexOf("(") - 1)) + "  (Config. with " + textBoxLastFile.Text + ")";

                ImportExport.DatasetToXml("config.xml", dataSetConfig);

                if (arcazeFound)
                {
                    timerstate = State.sendConfig;
                }
                else
                {
                    timerstate = State.init;
                }
                timerMain.Interval = timerInterval;
            }
        }

        private void ButtonPackage_Click(object sender, EventArgs e)
        {
            if (textBoxTestDataPackage.Text.Trim().Length > 0)
            {
                receivedData = textBoxTestDataPackage.Text;
            }
        }

        private void ButtonReadADC_Click(object sender, EventArgs e)
        {
            arcaze = ComboBoxArcaze.Text;
            ActivateArcaze(ref arcaze, false);

            int[] adcValue = new int[6];
            adcValue = ReadADC();

            int resolution = 1;

            trackBarADC1.Value = adcValue[0] / resolution;
            trackBarADC2.Value = adcValue[1] / resolution;
            trackBarADC3.Value = adcValue[2] / resolution;
            trackBarADC4.Value = adcValue[3] / resolution;
            trackBarADC5.Value = adcValue[4] / resolution;
            trackBarADC6.Value = adcValue[5] / resolution;
        }

        private void ButtonReset_Click(object sender, EventArgs e)
        {
            if (arcazeHid == null || !arcazeHid.Info.Connected)
            {
                return;
            }

            timerstate = State.reset;
        }

        private void ButtonSaveConfig_Click(object sender, EventArgs e)
        {
            SaveConfig();
            labelPleaseWait.Text = "Configuation saved.";
        }

        private void ButtonSaveToFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog()
            {
                InitialDirectory = Application.StartupPath,
                Filter = "D.A.C. Data Files *.xml|*.xml",
                Title = "Save a .xml File",
                FileName = textBoxLastFile.Text
            };
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxLastFile.Text = saveFileDialog1.FileName.Substring(saveFileDialog1.FileName.LastIndexOf('\\', saveFileDialog1.FileName.Length - 1) + 1);
                this.Text = this.Text.Substring(0, (this.Text.IndexOf("(") - 1)) + "  (Config. with " + textBoxLastFile.Text + ")";

                FillTypeClickable();
                FillTypeDCSID();
                ErasePulldown(); //------

                dataSetDisplaysLEDs.AcceptChanges();
                ImportExport.DatasetToXml(saveFileDialog1.FileName, dataSetDisplaysLEDs);
                SaveConfig();

                RefreshPulldown(); //------
                labelPleaseWait.Text = "All data saved.";
            }
        }

        private void ButtonSendPin_Click(object sender, EventArgs e)
        {
            arcazeDeviceIndex = GetArcazeInstanceFromLEDs(ComboBoxArcaze.Text);
            module = comboBoxModul.SelectedIndex;
            port = comboBoxConnector.SelectedIndex;
            pin = comboBoxPin.SelectedIndex;

            if (arcazeDeviceIndex > -1)
            {
                if (InitLEDDriver())
                {
                    switch (comboBoxModuleType.Text)
                    {
                        case "Display-Driver":
                            break;

                        case "LED-Driver 2": // ref module, ref port, ref pin, ref resolution,
                            resolution = 1;
                            arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port,
                                ref pin, ref resolution, Convert.ToDouble(checkBoxPinSetValue.Checked ? 1 : 0), 2, false, checkBoxLog.Checked);
                            break;

                        case "LED-Driver 3":
                            if (comboBoxLEDValue.Text == "")
                            {
                                comboBoxLEDValue.Text = "1.00";
                            }

                            resolution = 8;
                            pinString = comboBoxLEDValue.Text.Replace(",", ".");

                            arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port,
                                ref pin, ref resolution, double.Parse(pinString, CultureInfo.InvariantCulture), 3, false, checkBoxLog.Checked);
                            break;

                        case "Arcaze USB":
                            arcazeDevice[arcazeDeviceIndex].SetPinDirection(comboBoxConnector.SelectedIndex, comboBoxPin.SelectedIndex, 1); // direction 1 -> output

                            resolution = 1;
                            arcazeDevice[arcazeDeviceIndex].LEDWriteOutput(ref module, ref port,
                                ref pin, ref resolution, Convert.ToDouble(checkBoxPinSetValue.Checked ? 1 : 0), 4, false, checkBoxLog.Checked);
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

            if (stop && configID != -1)
            {
                UDP.UDPSender(textBoxIP.Text.Trim(), Convert.ToInt32(textBoxPortSender.Text), "S" + configID);
                ImportExport.LogMessage("Package processing stopped .. " + "S" + configID, true);
                configID = -1;
                SwitchAll(off);
                stop = true;
            }
            timerMain.Interval = timerInterval;
            timerstate = State.sendConfig;

            buttonStop.Text = stop ? "Start" : "Stop";
            buttonStop.ForeColor = (stop ? Color.Green : Color.Red);
        }

        private void ButtonWriteDigits_Click(object sender, EventArgs e)
        {
            string arcaze = ComboBoxArcaze.Text;
            ActivateArcaze(ref arcaze, false);
            arcazeDeviceIndex = GetArcazeInstanceFromDisplays(arcaze);

            if (arcazeDeviceIndex > -1)
            {
                digitString = "";

                for (int n = 0; n < textBoxDezValue.Text.Length; n++)
                {
                    characterValue = (byte)Convert.ToChar(textBoxDezValue.Text.Substring(n, 1));

                    if (characterValue >= 44 && characterValue <= 57 && characterValue != 47)
                    {
                        digitString += textBoxDezValue.Text.Substring(n, 1);
                    }
                }
                textBoxDezValue.Text = digitString;

                ConvertDezToSevenSegment(ref digit);

                arcazeDevice[arcazeDeviceIndex].InitDisplayDriver(int.Parse(comboBoxDevAddress.Text, NumberStyles.HexNumber), 0, trackBarDisplayBrightness.Value, 8);

                arcazeDevice[arcazeDeviceIndex].WriteDigitsToDisplayDriver(int.Parse(comboBoxDevAddress.Text, NumberStyles.HexNumber), ref digit,
                    int.Parse(textBoxDigitMask.Text, NumberStyles.HexNumber), checkBoxLog.Checked, checkBoxReverse.Checked, ref segmentIndex,
                    Convert.ToInt32(cbRefreshCycle.Text), Convert.ToInt32(cbDelayRefresh.Text));
            }
        }

        private void CheckBoxLog_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxLog.Checked)
            {
                logDetail = true;
            }
            else { logDetail = false; }
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
            {
                comboBoxConnector.SelectedIndex = 0;
            }
            else
            {
                comboBoxConnector.SelectedIndex = 2;
            }
        }

        private void ComboBoxArcaze_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                arcazeFromGrid = ComboBoxArcaze.Text;
                if (arcazeFromGrid != "") { ActivateArcaze(ref arcazeFromGrid, false); }
            }
            catch { return; }
        }

        private void ComboBoxArcaze_SelectedValueChanged(object sender, EventArgs e)
        {
            arcaze = ComboBoxArcaze.Text;

            if (arcaze != "")
            {
                ActivateArcaze(ref arcaze, false);
            }
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

        private void DataGridViewADC_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewArcaze_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewClickable_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewDisplays_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!arcazeFound)
            {
                return;
            }

            DataGridViewRow row = dataGridViewDisplays.Rows[e.RowIndex];

            if (e.ColumnIndex > 1 && e.ColumnIndex < 7 && e.ColumnIndex != 3)
            {
                buttonInit.Visible = true;
            }

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

        private void DataGridViewEncoderSend_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewJoystick_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //ImportExport.LogMessage("DataGridViewJoystick : " + e.ToString(), true);
        }

        private void DataGridViewEncoderValues_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridViewEncoderValues.Rows[e.RowIndex];

            if (row.Cells["arcaze_Encoder"].Value != null && row.Cells["Encoder_Number"].Value != null)
            {
                row.Cells["ArcazeAndEncoder"].Value = row.Cells["arcaze_Encoder"].Value.ToString() + " - " + row.Cells["Encoder_Number"].Value.ToString();
            }
            //dataGridViewEncoderSend.Refresh();
        }

        private void DataGridViewEncoderValues_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 2 && e.ColumnIndex < 7 && arcazeFound)
            {
                buttonInit.Visible = true;
            }
        }

        private void DataGridViewEncoderValues_DataError(object sender, DataGridViewDataErrorEventArgs e)
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
            {
                return;
            }

            DataGridViewRow row = dataGridViewLEDs.Rows[e.RowIndex];

            if (e.ColumnIndex > 3 && e.ColumnIndex < 11)
            {
                buttonInit.Visible = true;
            }

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

        private void DataGridViewKeystrokes_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 1 && e.ColumnIndex < 6 && arcazeFound)
            {
                buttonInit.Visible = true;
            }
        }

        private void DataGridViewKeystrokes_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void DataGridViewSwitches_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 1 && e.ColumnIndex < 9 && arcazeFound)
            {
                buttonInit.Visible = true;
            }
        }

        private void DataGridViewSwitches_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //ImportExport.LogMessage("dataGridViewSwitches : " + e.ToString(), true);
        }

        private void FormMain_Load(object sender, EventArgs e)
        {

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

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetLogs();

            if (tabControl1.SelectedIndex == 0) { labelCount.Text = (dataGridViewDisplays.RowCount - 1).ToString() + " records"; }
            if (tabControl1.SelectedIndex == 1) { labelCount.Text = (dataGridViewLEDs.RowCount - 1).ToString() + " records"; }
            if (tabControl1.SelectedIndex == 2) { labelCount.Text = (dataGridViewSwitches.RowCount - 1).ToString() + " records"; }
            if (tabControl1.SelectedIndex == 3) { labelCount.Text = (dataGridViewEncoderValues.RowCount - 1).ToString() + " records"; }
            if (tabControl1.SelectedIndex == 4) { labelCount.Text = (dataGridViewADC.RowCount - 1).ToString() + " records"; }
            if (tabControl1.SelectedIndex == 5) { labelCount.Text = (dataGridViewKeystrokes.RowCount - 1).ToString() + " records"; }
            if (tabControl1.SelectedIndex > 5) { labelCount.Text = ""; }

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
                arcaze = ComboBoxArcaze.Text;
                ActivateArcaze(ref arcaze, false);
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

        private void CheckBoxDebug_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxDebug.Checked)
            {
                debug = true;
                arcazeFound = true;
                //timerstate = State.initConfig;
                timerstate = State.startup;
            }
            else
            {
                debug = false;
                arcazeFound = false;
                timerstate = State.startup;
            }
            timerMain.Interval = timerInterval;
        }

        #endregion
    }
}
