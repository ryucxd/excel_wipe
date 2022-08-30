using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_wipe
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("MSACCESS.EXE");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill(); //kills each process :}
                    }
                    catch { }
                }
            }
        }
    }
}
