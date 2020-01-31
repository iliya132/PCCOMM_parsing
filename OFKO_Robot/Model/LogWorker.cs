using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OFKO_Robot.Model
{
    public static class LogWorker
    {

        public static void ClearPrevConsoleLine()
        {
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.Write("                                ");
            Console.SetCursorPosition(0, Console.CursorTop);
        }
    }
}
