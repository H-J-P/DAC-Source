#region Lizenz
//----------------------------------------------------------------------
//    Copyright (c) 2014 Heinz-Joerg Puhlmann
//
// Hiermit wird unentgeltlich, jeder Person, die eine Kopie der Software
// und der zugehörigen Dokumentationen (die "Software") erhält, die
// Erlaubnis erteilt, uneingeschränkt zu benutzen, inklusive und ohne
// Ausnahme, dem Recht, sie zu verwenden, kopieren, ändern, fusionieren,
// verlegen, verbreiten, unter-lizenzieren und/oder zu verkaufen, und 
// Personen, die diese Software erhalten, diese Rechte zu geben, unter
// den folgenden Bedingungen:
//
// Der obige Urheberrechtsvermerk und dieser Erlaubnisvermerk sind in 
// alle Kopien oder Teilkopien der Software beizulegen.
//
// DIE SOFTWARE WIRD OHNE JEDE AUSDRÜCKLICHE ODER IMPLIZIERTE GARANTIE 
// BEREITGESTELLT, EINSCHLIESSLICH DER GARANTIE ZUR BENUTZUNG FÜR DEN
// VORGESEHENEN ODER EINEM BESTIMMTEN ZWECK SOWIE JEGLICHER 
// RECHTSVERLETZUNG, JEDOCH NICHT DARAUF BESCHRÄNKT. IN KEINEM FALL SIND
// DIE AUTOREN ODER COPYRIGHTINHABER FÜR JEGLICHEN SCHADEN ODER SONSTIGE
// ANSPRUCH HAFTBAR ZU MACHEN, OB INFOLGE DER ERFÜLLUNG VON EINEM 
// VERTRAG, EINEM DELIKT ODER ANDERS IM ZUSAMMENHANG MIT DER BENUTZUNG
// ODER SONSTIGE VERWENDUNG DER SOFTWARE ENTSTANDEN
//----------------------------------------------------------------------
#endregion

using System;
//using Microsoft.DirectX.DirectInput;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DAC
{
    public static class Joystick
    {
        #region member
        //public static Boolean joystickActiv = false;
        //private static byte[] buttons = new byte[] { };
        //private static Device joystickDevice;
        //private static JoystickState state = new JoystickState();

        //private static List<Boolean[]> buttonPressed = new List<Boolean[]>();
        //private static List<DeviceCaps> caps = new List<DeviceCaps>();
        //private static List<Device> joysticks = new List<Device>();
        //private static List<Int32> joystickButtons = new List<Int32>();
        //public static List<String> joystickName = new List<String>();
        #endregion

        #region memberfunctions

        //public static void FindDevice(Form form)
        //{
        //    try
        //    {
        //        joystickActiv = false;
        //        DeviceList gameControllerList = Manager.GetDevices(DeviceClass.GameControl, EnumDevicesFlags.AttachedOnly);

        //        if (gameControllerList.Count > 0) // is there at least one device
        //        {
        //            for (int i = 1; i <= gameControllerList.Count; i++) // run trough devices
        //            {
        //                gameControllerList.MoveNext(); // choose next device
        //                DeviceInstance deviceInstance = (DeviceInstance)gameControllerList.Current; // create deviceinstance

        //                if (deviceInstance.DeviceType == DeviceType.Joystick) // is the selected device a joystick
        //                {
        //                    joystickDevice = new Device(deviceInstance.InstanceGuid); // create joystick device
        //                    joysticks.Add(joystickDevice);
        //                }
        //            }

        //            if (joysticks.Count > 0) // no joystick found
        //            {
        //                for (int n = 0; n < joysticks.Count; n++)
        //                {
        //                    joysticks[n].SetCooperativeLevel(form, CooperativeLevelFlags.Background | CooperativeLevelFlags.NonExclusive); // define interaction
        //                    joysticks[n].SetDataFormat(DeviceDataFormat.Joystick); // define that this device is a joystick
        //                    joysticks[n].Acquire(); // make it free

        //                    caps.Add(joysticks[n].Caps);
        //                    buttonPressed.Add(new Boolean[caps[n].NumberButtons]); // Flags for 'Pressed' Button

        //                    joystickButtons.Add(caps[n].NumberButtons);
        //                    joystickName.Add(joysticks[n].DeviceInformation.ProductName);
        //                }
        //                joystickActiv = true;

        //                if (joysticks.Count > 1)
        //                    ImportExport.LogMessage(joysticks.Count + " Joysticks found ... ", true);
        //                else
        //                    ImportExport.LogMessage(joysticks.Count + " Joystick found ... ", true);
        //            }
        //            else
        //                ImportExport.LogMessage("None joystick found ... ", true);
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        ImportExport.LogMessage("FindDevice ... " + e.ToString(), true);
        //    }
        //}

        //public static void UpdateJoystick(int joystick, ref string pressedButton)
        //{
        //    try
        //    {
        //        pressedButton = "";
        //        state = joysticks[joystick].CurrentJoystickState;
        //        buttons = state.GetButtons(); // Capture buttons state.

        //        for (int i = 0; i < buttonPressed[joystick].Length; i++)
        //        {
        //            if (buttons[i] != 0)
        //            {
        //                if (buttonPressed[joystick][i] == false)
        //                {
        //                    pressedButton += "Btn" + i + ":";
        //                    buttonPressed[joystick][i] = true;
        //                }
        //            }
        //            else
        //                buttonPressed[joystick][i] = false; // release button
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        ImportExport.LogMessage("UpdateJoystick ... " + e.ToString(), true);
        //    }
        //}

        #endregion
    }
}