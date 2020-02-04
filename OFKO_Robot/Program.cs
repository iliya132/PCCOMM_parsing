using System;
using System.Windows.Forms;

namespace OFKO_Robot
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            #region initialization
            Model.Equation eq = new Model.Equation();
            Interfaces.IExcelWorker excelWorker = new Model.ExcelWorker(eq);
            if (!eq.connected)
            {
                Console.WriteLine("Отсутствует соединение с Equation. Попробуйте разорвать связь и подключится снова. Нажмите Enter для завершения программы.");
                Console.ReadLine();
                return;
            }
            #endregion

            #region Диалоговое окно выбора файла
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Binary files (*.xlsb)|*.xlsb|Рабочая книга Excel (.xlsx)|.xlsx";
            if (openFileDialog.ShowDialog() != DialogResult.OK) { return; }
            #endregion

            #region Выполнение работы
            excelWorker.OpenFile(openFileDialog.FileName);
            excelWorker.Work();
            #endregion

            #region quit
            Console.WriteLine("Работа завершена. Нажмите enter для выхода.");
            Console.ReadLine();
            #endregion
        }
    }
}