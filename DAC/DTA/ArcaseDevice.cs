using System;
using System.Collections.Generic;
using System.Globalization;
using SimpleSolutions.Usb;

namespace DAC
{
    class ArcazeDevice
    {
        private ArcazeHid arcazeDevice;
        private List<byte> Digits = new List<byte>(8);
        UInt32 resolutionValue = 0;
        private string digitsValue = "";

        public string GetSerial
        {
            get
            {
                //return (this.arcazeDevice.Info.DeviceName + " " + this.arcazeDevice.Info.Serial);
                return (this.arcazeDevice.Info.Serial);
            }
        }

        public ArcazeDevice(DeviceInfo info)
        {
            this.arcazeDevice = new ArcazeHid();
            Connect(ref info);

            this.arcazeDevice.DeviceRemoved += new EventHandler<HidEventArgs>(this.DeviceRemoved);
            this.arcazeDevice.OurDeviceRemoved += new EventHandler<HidEventArgs>(this.OurDeviceRemoved);
        }

        public void Connect(ref DeviceInfo info)
        {
            this.arcazeDevice.Connect(info.Path);
        }

        public void Disconnect()
        {
            try
            {
                if (this.arcazeDevice.Info.Connected)
                    this.arcazeDevice.Disconnect();
            }
            catch
            { }
        }

        private void DeviceRemoved(object sender, HidEventArgs hidEventsArgs)
        {
            this.arcazeDevice.Disconnect();
        }

        private void OurDeviceRemoved(object sender, HidEventArgs hidEventsArgs)
        {
            this.arcazeDevice.Disconnect();
        }

        /// <summary>
        /// Init a display driver
        /// </summary>
        /// <param name="devAdress">The unique device address of the Display Module (set by the rotary switch on the board)</param>
        /// <param name="decodeMode">0x00 = No decoding / 0xFF = Code B Decoding (only lower data nibble used then), default = 0x00</param>
        /// <param name="intensity">0x00 ... 0x0F</param>
        /// <param name="scanLimit">4 ... 8 allowed (default = 8, no need to change)</param>
        public void InitDisplayDriver(int devAdress, int decodeMode, int intensity, int scanLimit)
        {
            try
            {
                ImportExport.LogMessage(GetSerial + " CmdMax7219DisplayInit(Modul: " + devAdress.ToString("X2") + ", decodeMode:" + decodeMode.ToString("X2") + ", intensity: "
                    + intensity.ToString("X2") + ", scanLimit: " + scanLimit.ToString("X2") + ")", true);

                this.arcazeDevice.Command.CmdMax7219DisplayInit(devAdress, decodeMode, intensity, scanLimit);
            }
            catch (Exception e)
            {
                ImportExport.LogMessage(GetSerial + " CmdMax7219DisplayInit(Modul: " + devAdress.ToString("X2") + ", decodeMode:" + decodeMode.ToString("X2") + ", intensity: "
                    + intensity.ToString("X2") + ", scanLimit: " + scanLimit.ToString("X2") + ") .. " + e.ToString(), true);
            }
        }

        public bool InitExtensionPort(ArcazeCommand.ExtModuleType modulType, int numModules, int bitsPerPort, int brightness)
        {
            try
            {
                if (brightness > 127)
                    brightness = 127;

                ImportExport.LogMessage(GetSerial + " CmdInitExtensionPort(" + modulType + " numModul: " +
                   numModules.ToString("X2") + ", Resolution: " + bitsPerPort.ToString("X2") +
                   ", Brightness: " + brightness.ToString("X2") + ")", true);

                this.arcazeDevice.Command.CmdInitExtensionPort(modulType, numModules, bitsPerPort, brightness);

                return true;
            }
            catch (Exception e)
            {
                ImportExport.LogMessage(GetSerial + " CmdInitExtensionPort(" + modulType + " numModul: " +
                    numModules.ToString("X2") + ", Resolution: " + bitsPerPort.ToString("X2") +
                    ", Brightness: " + brightness.ToString("X2") + ") ... " + e.ToString(), true);

                return false;
            }
        }

        public void LEDWriteOutput(ref int moduleNum, ref int connectorNum, ref int portNum, ref int resolution, Double data, int type, bool reverse, bool log)
        {
            if (arcazeDevice == null)
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

                resolutionValue = Convert.ToUInt32(ResolutionValue(ref resolution)); // LED-Driver 3
                data *= resolutionValue;

                if (data > resolutionValue)
                    data = resolutionValue;

                if (data < 0)
                    data = 0;
            }

            if (connectorNum > 1)
                connectorNum -= 2;

            try
            {
                if (log)
                    ImportExport.LogMessage(GetSerial + " WriteOutputPort(Modul: " + moduleNum.ToString("X2") + ", Connector: " + connectorNum.ToString("X2") + ", Pin: " + (portNum + 1).ToString("D2") + ", Value: " + (data == 0 ? "Off" : (type != 3 ? "On" : data.ToString())) + ")", true);

                this.arcazeDevice.Command.WriteOutputPort(moduleNum, connectorNum, portNum, Convert.ToUInt32(data), ArcazeCommand.OutputOperators.PlainWrite, false);
                this.arcazeDevice.Command.UpdateOutputPorts();
            }
            catch (Exception e)
            {
                ImportExport.LogMessage(GetSerial + " WriteOutputPort(Modul: " + moduleNum.ToString("X2") + ", Connector: " + connectorNum.ToString("X2") + ", Pin: " + (portNum + 1).ToString("D2") + ", Value: " + (data == 0 ? "Off" : (moduleNum == 0 ? "On" : data.ToString())) + ") ... " + e.ToString(), true);
            }
        }

        private uint ResolutionValue(ref int resolution)
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

        /// <summary>
        /// Set the pin direction on the arcaze
        /// </summary>
        /// <param name="port">Connector A = 0;B = 1;C = 2</param>
        /// <param name="pin">0 - 19</param>
        /// <param name="direction">0 = input; 1 = output</param>
        public void SetPinDirection(int port, int pin, int direction)
        {
            try
            {
                ImportExport.LogMessage(GetSerial + " SetPinDirection(Connector: " + port.ToString("X2") + ", Pin: " + (pin + 1).ToString("D2") + ", Direction: " + (direction == 1 ? "Output" : "Input") + ")", true);

                this.arcazeDevice.Command.CmdSetPinDirection(port, pin, direction);
            }
            catch (Exception f)
            {
                ImportExport.LogMessage(GetSerial + " SetPinDirection(Connector: " + port.ToString("X2") + ", Pin: " + (pin + 1).ToString("D2") + ", Direction: " + (direction == 1 ? "Output" : "Input") + ") .. " + f.ToString(), true);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="devAdress">0 .. 15</param>
        /// <param name="digit">SevenSegment Value</param>
        /// <param name="digitMask">0x00 .. 0xFF</param>
        public void WriteDigitsToDisplayDriver(int devAdress, ref string[] digit, int digitMask, bool log, bool reverse, ref int segmentIndex, int cycle, int delay)
        {
            Digits = new List<byte>(8);

            if (reverse)
            {
                for (int n = 0; n < 8; n++)
                {
                    if (digit[n] == "0")
                        Digits.Add(0);
                    else
                        break;
                }
                for (int n = segmentIndex; n > 0; n--)
                    Digits.Add(byte.Parse(digit[n - 1], NumberStyles.HexNumber));
            }
            else
            {
                for (int n = 0; n < 8; n++)
                    Digits.Add(byte.Parse(digit[n], NumberStyles.HexNumber));
            }
            try
            {
                arcazeDevice.Command.CmdMax7219WriteDigits(devAdress, Digits, digitMask);

                for (int n = 0; n < cycle; n++)
                {
                    StopWatch.NOP(delay);
                    arcazeDevice.Command.CmdMax7219WriteDigits(devAdress, Digits, digitMask);
                }
                if (log)
                {
                    digitsValue = "";

                    for (int n = 7; n > -1; n--)
                        digitsValue += Digits[n].ToString("X2") + " ";

                    ImportExport.LogMessage(GetSerial + " CmdMax7219WriteDigits(Modul: " + devAdress.ToString("X2") + ", Digits: " + digitsValue + ", Mask: " + (digitMask).ToString("X2") + ")", true);
                }
            }
            catch (Exception e)
            {
                digitsValue = "";

                for (int n = 0; n < 8; n++)
                    digitsValue = Digits[n].ToString("X2") + " ";

                ImportExport.LogMessage(GetSerial + " CmdMax7219WriteDigits(Modul: " + devAdress.ToString("X2") + ", Digits: " + digitsValue + ", Mask: " + (digitMask).ToString("X2")
                     + ") .. " + e.ToString(), true);
            }
        }
    }
}
