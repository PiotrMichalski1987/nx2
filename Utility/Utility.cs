using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WrkWebApp.Utility
{
    public static class Utility
    {


        public static string LongCodeToShortCode(string longCode)
        {
            string newCode = "";

            newCode += longCode[6];
            newCode += longCode[7];
            newCode += longCode[8];

            return newCode;

        }

        public static string SixDigCodeToShortCode(string longCode)
        {
            string newCode = "";

            newCode += longCode[2];
            newCode += longCode[3];
            newCode += longCode[4];

            return newCode;

        }

        


    }
}