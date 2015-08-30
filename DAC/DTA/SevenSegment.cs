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
using System.Globalization;

namespace DAC
{
    class SevenSegment
    {
        //      dp a b c d e f g
        //  bit: 7,6,5,4,3,2,1,0            hex
        //     { 0,1,1,1,1,1,1,0 }  // = 0  0x7E
        //     { 0,0,1,1,0,0,0,0 }  // = 1  0x30
        //     { 0,1,1,0,1,1,0,1 }  // = 2  0x6D
        //     { 0,1,1,1,1,0,0,1 }  // = 3  0x79
        //     { 0,0,1,1,0,0,1,1 }  // = 4  0x33
        //     { 0,1,0,1,1,0,1,1 }  // = 5  0x5B
        //     { 0,1,0,1,1,1,1,1 }  // = 6  0x5F
        //     { 0,1,1,1,0,0,0,0 }  // = 7  0x70
        //     { 0,1,1,1,1,1,1,1 }  // = 8  0x7F
        //     { 0,1,1,1,1,0,1,1 }  // = 9  0x7B
        //     { 1,0,0,0,0,0,0,0 }  // = .  0x80

        private enum valuePattern
        {
            None = 0x0, Zero = 0x7E, One = 0x30, Two = 0x6D, Three = 0x79,
            Four = 0x33, Five = 0x5B, Six = 0x5F, Seven = 0x70,
            Eight = 0x7F, Nine = 0x7B
        }

        private static string customPattern = "";

        public static string GetNonePattern()
        {
            return ((int)valuePattern.None).ToString("X");
        }

        /// <summary>
        /// Character to be displayed on the seven segments. Supported characters
        /// are digits and most letters.
        /// </summary>
        public static string GetValuePattern(int value, bool dezimalpoint)
        {
            customPattern = "";

            try
            {
                int tempValue = Convert.ToInt32(value);

                if (tempValue > 9) tempValue = 9;
                if (tempValue < 0) tempValue = 0;

                switch (tempValue)
                {
                    case 0: customPattern = ((int)valuePattern.Zero).ToString("X"); break;
                    case 1: customPattern = ((int)valuePattern.One).ToString("X"); break;
                    case 2: customPattern = ((int)valuePattern.Two).ToString("X"); break;
                    case 3: customPattern = ((int)valuePattern.Three).ToString("X"); break;
                    case 4: customPattern = ((int)valuePattern.Four).ToString("X"); break;
                    case 5: customPattern = ((int)valuePattern.Five).ToString("X"); break;
                    case 6: customPattern = ((int)valuePattern.Six).ToString("X"); break;
                    case 7: customPattern = ((int)valuePattern.Seven).ToString("X"); break;
                    case 8: customPattern = ((int)valuePattern.Eight).ToString("X"); break;
                    case 9: customPattern = ((int)valuePattern.Nine).ToString("X"); break;
                }
                if (dezimalpoint)
                    customPattern = (int.Parse(customPattern, NumberStyles.HexNumber) + 128).ToString("X");
            }
            catch { customPattern = ((int)valuePattern.None).ToString("X"); }
            return customPattern;
        }
    }
}
