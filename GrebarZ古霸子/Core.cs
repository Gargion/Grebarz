using System;
using System.Collections.Generic;
using System.Text;

namespace Grebarz古霸子
{
    public class Core
    {
        private static readonly string SYSTEM_NAME = "【Grebarz】"; // 程序名稱
        private static readonly string SYSTEM_VERSION = "beta 0.0.1";  // 程序版本號
        private static readonly string PRODUCT_SITE = "https://gz-rebarplanner.com/";

        public static string Gz_GetSystemName()
        {
            return SYSTEM_NAME;
        }
        public static string Gz_GetSystemVersion()
        {
            return SYSTEM_VERSION;
        }

        public static string Gz_GetProductSite()
        {
            return PRODUCT_SITE;
        }
    }
}
