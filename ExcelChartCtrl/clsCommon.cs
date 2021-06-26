

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ctrl_Dll;

namespace ExcelChartCtrl
{

    public struct structDir
    {
        public static string App;
        public static string DeskTop;
        public static string Document;
    }



    class clsCommon
    {

        public static cls_FileCtrl g_clsFC = new cls_FileCtrl();
        public static string strExcelFileName;

        public static void SetDir()
        { 
            structDir.App = g_clsFC.App_Directory_Acquisition();
            structDir.DeskTop = g_clsFC.Desk_Top_Directory();
            structDir.Document = g_clsFC.Mydocument_Directory();

            strExcelFileName = structDir.App + clsInit.g_EXCEL_FILE_NAME;
        }
    }
}
