using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAC
{
    //-------------------------------------------------------------------------------------------------------------------
    // AB AB AB AB AB AB AB AB AB
    // 00 01 11 10 00 01 11 10 00       Note: Only 1 bit will change at a time.
    // clockwise-->  <--anticlockwise
    //-------------------------------------------------------------------------------------------------------------------
    // By storing the former state of the two bits you can determine which direction the encoder is turned.
    // When 'oldValue' equals 00 and the new value is 01 it's rotating clockwise.
    //-------------------------------------------------------------------------------------------------------------------
    class RotaryEncoder
    {
        static int newValue = 0;
        static int retValue = 0;
        static bool flipValue = true;

        public static int Encoder(int pinA, int pinB, ref int oldValue)
        {
            newValue = (pinA * 2) + pinB;

            if (newValue == oldValue)
                return 0;

            flipValue = !flipValue;

            if (flipValue)
            {
                oldValue = newValue;
                return 0;
            }

            switch (oldValue)
            {
                case 0:
                    if (newValue == 1) retValue = 1;    // clockwise
                    if (newValue == 2) retValue = -1;   // anticlockwise
                    break;

                case 1:
                    if (newValue == 3) retValue = 1;
                    if (newValue == 0) retValue = -1;
                    break;

                case 2:
                    if (newValue == 0) retValue = 1;
                    if (newValue == 3) retValue = -1;
                    break;

                case 3:
                    if (newValue == 2) retValue = 1;
                    if (newValue == 1) retValue = -1;
                    break;
            }

            oldValue = newValue;

            return retValue;
        }

        public static int Encoder(ref int newValue, ref int oldValue)
        {
            if (newValue == oldValue)
                return 0;

            switch (oldValue)
            {
                case 0:
                    if (newValue == 1) retValue = 1;    // clockwise
                    if (newValue == 2) retValue = -1;   // anticlockwise
                    break;

                case 1:
                    if (newValue == 3) retValue = 1;
                    if (newValue == 0) retValue = -1;
                    break;

                case 2:
                    if (newValue == 0) retValue = 1;
                    if (newValue == 3) retValue = -1;
                    break;

                case 3:
                    if (newValue == 2) retValue = 1;
                    if (newValue == 1) retValue = -1;
                    break;
            }
            return retValue;
        }
    }
}
