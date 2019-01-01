using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCuiHuaJi
{
    class Conversions
    {
        // Standard Conversion Functions
        [Flags]
        internal enum XlType : int
        {
            XlTypeNumber = 0x0001,
            XlTypeString = 0x0002,
            XlTypeBoolean = 0x0004,
            XlTypeReference = 0x0008,
            XlTypeError = 0x0010,
            XlTypeArray = 0x0040,
            XlTypeMissing = 0x0080,
            XlTypeEmpty = 0x0100,
            XlTypeInt = 0x0800,     // int16 in XlOper, int32 in XlOper12, never passed into UDF
        }


        public static double dnaConvertToDouble(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeNumber);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
            {
                return (double)result;
            }

            // We give up.
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to Int32.");
        }

        
        public static string dnaConvertToString(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeString);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
            {
                return (string)result;
            }

            // Not sure how this can happen...
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to String.");
        }


        public static DateTime dnaConvertToDateTime(object value)
        {
            try
            {
                return DateTime.FromOADate(dnaConvertToDouble(value));
            }
            catch
            {
                // Might exceed range of DateTime
                throw new InvalidCastException("Value " + value.ToString() + " could not be converted to DateTime.");
            }
        }


        public static bool dnaConvertToBoolean(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeBoolean);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
                return (bool)result;

            // failed - as a fallback, try to convert to a double
            retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeNumber);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
                return ((double)result != 0.0);

            // We give up.
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to Boolean.");
        }


        public static int dnaConvertToInt32(object value)
        {
            return checked((int)dnaConvertToInt64(value));
        }


        public static short dnaConvertToInt16(object value)
        {
            return checked((short)dnaConvertToInt64(value));
        }


        public static ushort dnaConvertToUInt16(object value)
        {
            return checked((ushort)dnaConvertToInt64(value));
        }


        public static decimal dnaConvertToDecimal(object value)
        {
            return checked((decimal)dnaConvertToDouble(value));
        }


        public static long dnaConvertToInt64(object value)
        {
            return checked((long)Math.Round(dnaConvertToDouble(value), MidpointRounding.ToEven));
        }
    }
}
