using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OFKO_Robot.Interfaces
{
    interface IExcelWorker
    {
        void CreateFile(string filePath);
        void OpenFile(string filePath);
        void SaveFile(string filePath);
        void Save();
        void Write(int x, int y, string text);
        string Read(int x, int y);
        void Fill(int x, int y, Color color);
        void Work();
        void Dispose();
    }
}
