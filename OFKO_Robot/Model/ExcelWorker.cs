using System;
using System.Drawing;
using LinqToExcel;
using OfficeOpenXml;
using System.Linq;
using System.IO;
using System.Collections.Generic;

namespace OFKO_Robot.Model
{
    class ExcelWorker : Interfaces.IExcelWorker, IDisposable
    {
        ExcelPackage excel;
        ExcelWorksheet sheet;
        private readonly Equation equation;

        public ExcelWorker(Equation equation)
        {
            this.equation = equation;
        }
        public void CreateFile(string filePath)
        {
            excel = new ExcelPackage(new FileInfo(filePath));
            sheet = excel.Workbook.Worksheets.Add("Sheet1");
        }
        public void Dispose()
        {
            if (IsOn())
            {
                sheet.Dispose();
                excel.Dispose();
            }
        }
        public void Fill(int x, int y, Color color)
        {
            if (IsOn())
            {
                sheet.Cells[x, y].Style.Fill.BackgroundColor.SetColor(color);
            }
            else
            {
                throw new NullReferenceException("Для работы метода Fill необходим открытый экземпляр excel");
            }
        }
        /// <summary>
        /// Opens xlsxFile
        /// </summary>
        /// <param name="filePath"></param>
        public void OpenFile(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            //xlsb разбирается с помощью linq to Excel
            if (file.Extension.Equals(".xlsb"))
            {
                string newFileName = $"{file.DirectoryName}\\Done_{DateTime.Now.ToString("dd.MM.yyyyss")}.xlsx";
                CreateFile(newFileName);
                ExcelQueryFactory excelFactory = new ExcelQueryFactory(filePath);
                IQueryable<RowNoHeader> data = from i in excelFactory.WorksheetNoHeader(0) select i;
                RowNoHeader[] rows = data.ToArray();
                for (int i = 0; i < rows.Length; i++)
                {
                    Write(i + 2, 1, rows[i][0].Value.ToString());
                }
                //xlsx разбираем Epplus
            }
            else if (file.Extension.Equals(".xlsx"))
            {
                excel = new ExcelPackage(file);
                sheet = excel.Workbook.Worksheets[1];
            }
        }
        private void ClearPrevConsoleLine()
        {
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.Write("                                ");
            Console.SetCursorPosition(0, Console.CursorTop);
        }
        public void Work()
        {
            placeHeaders();
            List<FL_data> mans = new List<FL_data>();
            Console.WriteLine("Reading data from Excel (Persons)");
            double onePercent = 100.00 / sheet.Dimension.Rows;
            double progress = 0;
            for (int i = 2; i <= sheet.Dimension.Rows; i++)
            {
                Console.WriteLine($"Done: {progress:F2}%");
                mans.Add(equation.ReadData(sheet.Cells[i, 1].Text));
                progress = onePercent * i;
                ClearPrevConsoleLine();
            }
            Console.WriteLine("Getting AccountsData");
            equation.fillAccounts(mans);
            Console.WriteLine("Done.");
            Console.WriteLine("Exporting data to .xlsx");
            int rowIndex = 2;
            for (int i = 0; i < mans.Count; i++)
            {
                for (int j = 0; j < mans[i].Accounts.Count; j++)
                {
                    Write(rowIndex, 1, mans[i].Pin);
                    Write(rowIndex, 2, mans[i].Type);
                    Write(rowIndex, 3, mans[i].Identification);
                    Write(rowIndex, 4, mans[i].ApproveType);
                    Write(rowIndex, 5, mans[i].FIO);
                    Write(rowIndex, 6, NormalizeData(mans[i].BirthDate));
                    Write(rowIndex, 7, mans[i].mnemonicFull);
                    Write(rowIndex, 8, mans[i].mnemonicShort);
                    Write(rowIndex, 9, mans[i].regCountry);
                    Write(rowIndex, 10, mans[i].Inn);
                    Write(rowIndex, 11, mans[i].SNILS);
                    Write(rowIndex, 12, mans[i].IsClient ? "Да" : "Нет");
                    Write(rowIndex, 13, NormalizeData(mans[i].BecomeContragentDate));
                    Write(rowIndex, 14, NormalizeData(mans[i].BecomeClientDate));
                    Write(rowIndex, 15, mans[i].DocTypeInt);
                    Write(rowIndex, 16, mans[i].DocTypeStr);
                    Write(rowIndex, 17, mans[i].DocSeries);
                    Write(rowIndex, 18, mans[i].DocNumber);
                    Write(rowIndex, 19, mans[i].DocGiver);
                    Write(rowIndex, 20, mans[i].DocKPPCode);
                    Write(rowIndex, 21, mans[i].DocResponsibleUnit);
                    Write(rowIndex, 22, mans[i].Citizenship);
                    Write(rowIndex, 23, mans[i].WorkPhone);
                    Write(rowIndex, 24, mans[i].MobilePhone);
                    Write(rowIndex, 25, mans[i].Email);
                    Write(rowIndex, 26, mans[i].Reserve1);
                    Write(rowIndex, 27, mans[i].Reserve2);
                    Write(rowIndex, 28, mans[i].AnyBeneficiary ? "Да" : "Нет");
                    Write(rowIndex, 29, mans[i].PDL ? "Да" : "Нет");
                    Write(rowIndex, 30, mans[i].BeneficiaryOwner ? "Да" : "Нет");
                    Write(rowIndex, 31, mans[i].Reputation);
                    Write(rowIndex, 32, mans[i].FinancialPosition);
                    Write(rowIndex, 33, NormalizeData(mans[i].ActualizationDate));
                    Write(rowIndex, 34, mans[i].ActualizationEmployee);
                    Write(rowIndex, 35, mans[i].UpdateType);
                    Write(rowIndex, 36, NormalizeData(mans[i].LastUpdatedDate));
                    Write(rowIndex, 37, mans[i].UpdateEmployee);
                    Write(rowIndex, 38, mans[i].IsCrossedTheBorder ? "Да" : "Нет");
                    Write(rowIndex, 39, NormalizeData(mans[i].CrossBorderDate));
                    Write(rowIndex, 40, mans[i].ResidenceDocTypeInt);
                    Write(rowIndex, 41, mans[i].ResidenceDocTypeStr);
                    Write(rowIndex, 42, mans[i].ResidenceSerieAndNumber);
                    Write(rowIndex, 43, NormalizeData(mans[i].ResidenceDateStart));
                    Write(rowIndex, 44, NormalizeData(mans[i].ResidenceDateEnd));
                    Write(rowIndex, 45, mans[i].MigrationCardTypeInt);
                    Write(rowIndex, 46, mans[i].MigrationCardNumber);
                    Write(rowIndex, 47, NormalizeData(mans[i].MigrationCardDateStart));
                    Write(rowIndex, 48, NormalizeData(mans[i].MigrationCardDateEnd));
                    Write(rowIndex, 49, mans[i].RiskLevelInt.ToString());
                    Write(rowIndex, 50, mans[i].RiskLevelStr);
                    Write(rowIndex, 51, mans[i].Accounts[j].Number);
                    Write(rowIndex, 52, mans[i].Accounts[j].Type);
                    Write(rowIndex, 53, mans[i].Accounts[j].OpenDate.ToString("dd.MM.yyyy"));
                    Write(rowIndex, 54, mans[i].Accounts[j].Mnemonic);
                    Write(rowIndex, 55, mans[i].Accounts[j].MnemonicFull);
                    Write(rowIndex, 56, mans[i].Accounts[j].EmployeeOpenedFIO);
                    Write(rowIndex, 57, mans[i].Accounts[j].AccountingMode);
                    Write(rowIndex, 59, mans[i].RegistrationAddress);
                    Write(rowIndex, 60, mans[i].RegistrationAddressFull);
                    Write(rowIndex, 61, NormalizeData(mans[i].RegistrationUpdateDate));
                    Write(rowIndex, 62, mans[i].RegistrationPhone);
                    Write(rowIndex, 63, mans[i].FactAddress);
                    Write(rowIndex, 64, NormalizeData(mans[i].FactAddressUpdateDate));
                    Write(rowIndex, 65, mans[i].FactAddressPhone);
                    Write(rowIndex, 66, mans[i].RecordStatus);
                    Write(rowIndex, 67, mans[i].BirthPlace);
                    Write(rowIndex, 68, mans[i].BirthPlaceCountryCode);
                    Write(rowIndex, 69, mans[i].BirthPlaceCountry);
                    Write(rowIndex, 70, mans[i].BirthPlaceRegionCode);
                    Write(rowIndex, 71, mans[i].BirthPlaceRegionType);
                    Write(rowIndex, 72, mans[i].BirthPlaceRegion);
                    Write(rowIndex, 73, mans[i].BirthPlaceCityType);
                    Write(rowIndex, 74, mans[i].BirthPlaceCity);
                    Write(rowIndex, 75, mans[i].BirthPlaceLocalityType);
                    Write(rowIndex, 76, mans[i].BirthPlaceLocality);
                    rowIndex++;
                }
            }
            
            Save();
            equation.CloseConnection();

        }

        private string NormalizeData(DateTime dateTime)
        {
            return dateTime.Year < 1000 ? string.Empty : dateTime.ToString("dd.MM.yyyy");
        }
        /// <summary>
        /// Разместить заголовки столбцов
        /// </summary>
        private void placeHeaders()
        {
            Write(1, 1, "Пин");
            Write(1, 2, "Тип");
            Write(1, 3, "Идентификация");
            Write(1, 4, "Тип подтверждения");
            Write(1, 5, "ФИО");
            Write(1, 6, "Дата рождения");
            Write(1, 7, "Мнемоника_расшифровка");
            Write(1, 8, "Мнемоника");
            Write(1, 9, "Страна регистрации");
            Write(1, 10, "ИНН");
            Write(1, 11, "СНИЛС");
            Write(1, 12, "Клиент?");
            Write(1, 13, "Дата перехода в контрагенты");
            Write(1, 14, "Дата перехода в клиенты");
            Write(1, 15, "Тип документа(цифр.)");
            Write(1, 16, "Тип документа(строка)");
            Write(1, 17, "Серия");
            Write(1, 18, "Номер");
            Write(1, 19, "Документ выдан");
            Write(1, 20, "КПП");
            Write(1, 21, "Ответственное подразделение");
            Write(1, 22, "Гражданство");
            Write(1, 23, "Рабочий телефон");
            Write(1, 24, "Мобильный телефон");
            Write(1, 25, "Email");
            Write(1, 26, "Резерв1");
            Write(1, 27, "Резерв2");
            Write(1, 28, "Наличие выгодоприобретателя");
            Write(1, 29, "Принадлежит к ПДЛ");
            Write(1, 30, "Наличие бенефициарного владельца");
            Write(1, 31, "Репутация");
            Write(1, 32, "Финансовое положение");
            Write(1, 33, "Дата актуализации");
            Write(1, 34, "Сотрудник, проводивший актуализацию");
            Write(1, 35, "Тип обновления");
            Write(1, 36, "Дата последнего обновления");
            Write(1, 37, "Сотрудник, производивший обновление");
            Write(1, 38, "Пересекал границу");
            Write(1, 39, "Дата пересечения границы");
            Write(1, 40, "Код типа документа, разрешающего пребывание");
            Write(1, 41, "Расшифровка документа,");
            Write(1, 42, "Серия и номер");
            Write(1, 43, "Дата начала");
            Write(1, 44, "Дата окончания");
            Write(1, 45, "Тип МК");
            Write(1, 46, "Номер МК");
            Write(1, 47, "Дата начала МК");
            Write(1, 48, "Дата окончания МК");
            Write(1, 49, "Код уровня риска");
            Write(1, 50, "Расшифровка уровня риска");
            Write(1, 51, "Номер счета");
            Write(1, 52, "Тип счета");
            Write(1, 53, "Дата открытия счета");
            Write(1, 54, "Отделение счета");
            Write(1, 55, "Расширенное отделение счета");
            Write(1, 56, "Сотрудник, открывший счет");
            Write(1, 57, "Бух.режим");
            Write(1, 58, "Дата закрытия счета");
            Write(1, 59, "Адрес регистрации");
            Write(1, 60, "Адрес регистрации полный");
            Write(1, 61, "Дата обновления адреса регистрации");
            Write(1, 62, "RegistrationPhone");
            Write(1, 63, "Фактический адрес");
            Write(1, 64, "Дата обновления факт. адреса");
            Write(1, 65, "FactAddressPhone");
            Write(1, 66, "Статус записи");
            Write(1, 67, "Место рождения");
            Write(1, 68, "Код страны рождения");
            Write(1, 69, "Страна рождения");
            Write(1, 70, "Код региона рождения");
            Write(1, 71, "Тип региона рождения");
            Write(1, 72, "Регион рождения");
            Write(1, 73, "Тип города рождения");
            Write(1, 74, "Город рождения");
            Write(1, 75, "Тип населенного пункта рождения");
            Write(1, 76, "Населенный пункт рождения");
        }
        public string Read(int x, int y)
        {
            if (IsOn())
            {
                return sheet.Cells[x, y].Text;
            }
            else
            {
                throw new NullReferenceException("Для чтения данных необходим открытый экземпляр Excel");
            }
        }
        public void SaveFile(string filePath)
        {
            if (IsOn())
            {
                excel.SaveAs(new System.IO.FileInfo(filePath));
            }
            else
            {
                throw new NullReferenceException("Для сохранения файла необходим открытый экземпляр Excel");
            }
        }
        public void Save()
        {
            if (IsOn())
            {
                excel.Save();
            }
            else
            {
                throw new NullReferenceException("Для сохранения файла необходим открытый экземпляр Excel");
            }
        }
        public void Write(int x, int y, string text)
        {
            if (IsOn())
            {
                sheet.Cells[x, y].Value = text;
            }
            else
            {
                throw new NullReferenceException("Для работы метода Write необходим открытый экземпляр excel");
            }
        }
        private bool IsOn()
        {
            if (excel == null || sheet == null)
                return false;
            return true;
        }
    }
}
