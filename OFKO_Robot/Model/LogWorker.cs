using System;

namespace OFKO_Robot.Model
{
    /// <summary>
    /// Класс для работы с консолью.
    /// </summary>
    public static class LogWorker
    {
        /// <summary>
        /// Очистить предыдущую строку
        /// </summary>
        public static void ClearPrevConsoleLine()
        {
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.Write("                                ");
            Console.SetCursorPosition(0, Console.CursorTop);
        }
    }
}
