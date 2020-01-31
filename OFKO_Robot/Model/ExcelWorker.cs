using System;
using System.Drawing;
using LinqToExcel;
using OfficeOpenXml;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;

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
                string newFileName = $"{file.DirectoryName}\\Done_{DateTime.Now.ToString("dd.MM.yyyyss.fff")}.xlsx";
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
        public void Work()
        {
            placeHeaders();
            List<FL_data> mans = new List<FL_data>();
            Console.WriteLine("Reading data from Excel (Persons)");
            double onePercent = 100.00 / sheet.Dimension.Rows;
            double progress = 0;
            Statistic statistic = new Statistic(3);

            for (int i = 2; i <= sheet.Dimension.Rows; i++)
            {
                Application.DoEvents();
                Console.WriteLine($"Done: {progress:F2}% ({i - 2}/{sheet.Dimension.Rows - 2})");
                mans.Add(equation.ReadData(sheet.Cells[i, 1].Text));
                progress = onePercent * i;
                LogWorker.ClearPrevConsoleLine();
                statistic.OperationDone();
            }
            Console.WriteLine("Done.");
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
                    Write(rowIndex, 7, mans[i].BirthPlace);
                    Write(rowIndex, 8, mans[i].mnemonicFull);
                    Write(rowIndex, 9, mans[i].mnemonicShort);
                    Write(rowIndex, 10, mans[i].regCountry);
                    Write(rowIndex, 11, mans[i].Inn);
                    Write(rowIndex, 12, mans[i].SNILS);
                    Write(rowIndex, 13, mans[i].IsClient);
                    Write(rowIndex, 14, NormalizeData(mans[i].BecomeContragentDate));
                    Write(rowIndex, 15, NormalizeData(mans[i].BecomeClientDate));
                    Write(rowIndex, 16, mans[i].DocTypeInt);
                    Write(rowIndex, 17, mans[i].DocTypeStr);
                    Write(rowIndex, 18, mans[i].DocSeries);
                    Write(rowIndex, 19, mans[i].DocNumber);
                    Write(rowIndex, 20, mans[i].DocGiver);
                    Write(rowIndex, 21, mans[i].DocKPPCode);
                    Write(rowIndex, 22, NormalizeData(mans[i].DocOpenDate));
                    Write(rowIndex, 23, NormalizeData(mans[i].DocClosedDate));
                    Write(rowIndex, 24, mans[i].DocResponsibleUnit);
                    Write(rowIndex, 25, mans[i].Citizenship);
                    Write(rowIndex, 26, mans[i].WorkPhone);
                    Write(rowIndex, 27, mans[i].MobilePhone);
                    Write(rowIndex, 28, mans[i].Email);
                    Write(rowIndex, 29, mans[i].Reserve1);
                    Write(rowIndex, 30, mans[i].Reserve2);
                    Write(rowIndex, 31, mans[i].AnyBeneficiary);
                    Write(rowIndex, 32, mans[i].PDL);
                    Write(rowIndex, 33, mans[i].BeneficiaryOwner);
                    Write(rowIndex, 34, mans[i].Reputation);
                    Write(rowIndex, 35, mans[i].FinancialPosition);
                    Write(rowIndex, 36, NormalizeData(mans[i].ActualizationDate));
                    Write(rowIndex, 37, mans[i].ActualizationEmployee);
                    Write(rowIndex, 38, mans[i].UpdateType);
                    Write(rowIndex, 39, NormalizeData(mans[i].LastUpdatedDate));
                    Write(rowIndex, 40, mans[i].UpdateEmployee);
                    Write(rowIndex, 41, mans[i].IsCrossedTheBorder);
                    Write(rowIndex, 42, NormalizeData(mans[i].CrossBorderDate));
                    Write(rowIndex, 43, mans[i].ResidenceDocTypeInt);
                    Write(rowIndex, 44, mans[i].ResidenceDocTypeStr);
                    Write(rowIndex, 45, mans[i].ResidenceSerieAndNumber);
                    Write(rowIndex, 46, NormalizeData(mans[i].ResidenceDateStart));
                    Write(rowIndex, 47, NormalizeData(mans[i].ResidenceDateEnd));
                    Write(rowIndex, 48, mans[i].MigrationCardTypeInt);
                    Write(rowIndex, 49, mans[i].MigrationCardNumber);
                    Write(rowIndex, 50, NormalizeData(mans[i].MigrationCardDateStart));
                    Write(rowIndex, 51, NormalizeData(mans[i].MigrationCardDateEnd));
                    Write(rowIndex, 52, mans[i].RiskLevelInt.ToString());
                    Write(rowIndex, 53, mans[i].RiskLevelStr);
                    Write(rowIndex, 54, mans[i].Accounts[j].Number);
                    Write(rowIndex, 55, mans[i].Accounts[j].Type);
                    Write(rowIndex, 56, mans[i].Accounts[j].OpenDate.ToString("dd.MM.yyyy"));
                    Write(rowIndex, 57, mans[i].Accounts[j].Mnemonic);
                    Write(rowIndex, 58, mans[i].Accounts[j].MnemonicFull);
                    Write(rowIndex, 59, mans[i].Accounts[j].EmployeeOpenedFIO);
                    Write(rowIndex, 60, mans[i].Accounts[j].AccountingMode);
                    Write(rowIndex, 61, NormalizeData(mans[i].Accounts[j].CloseDate));
                    Write(rowIndex, 62, mans[i].RegistrationAddress);
                    Write(rowIndex, 63, mans[i].PrimePhone);
                    Write(rowIndex, 64, mans[i].RegistrationAddressFull);
                    Write(rowIndex, 65, NormalizeData(mans[i].RegistrationUpdateDate));
                    Write(rowIndex, 66, mans[i].RegistrationPhone);
                    Write(rowIndex, 67, mans[i].FactAddress);
                    Write(rowIndex, 68, NormalizeData(mans[i].FactAddressUpdateDate));
                    Write(rowIndex, 69, mans[i].FactAddressPhone);
                    Write(rowIndex, 70, mans[i].RecordStatus);
                    Write(rowIndex, 71, mans[i].BirthPlace);
                    Write(rowIndex, 72, mans[i].BirthPlaceCountryCode);
                    Write(rowIndex, 73, mans[i].BirthPlaceCountry);
                    Write(rowIndex, 74, mans[i].BirthPlaceRegionCode);
                    Write(rowIndex, 75, mans[i].BirthPlaceRegionType);
                    Write(rowIndex, 76, mans[i].BirthPlaceRegion);
                    Write(rowIndex, 77, mans[i].BirthPlaceCityType);
                    Write(rowIndex, 78, mans[i].BirthPlaceCity);
                    Write(rowIndex, 79, mans[i].BirthPlaceLocalityType);
                    Write(rowIndex, 80, mans[i].BirthPlaceLocality);
                    Write(rowIndex, 81, mans[i].BirthPlaceFull);

                    rowIndex++;
                }
            }
            statistic.Commit();
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
            Write(1, 7, "Место рождения");
            Write(1, 8, "Мнемоника_расшифровка");
            Write(1, 9, "Мнемоника");
            Write(1, 10, "Страна регистрации");
            Write(1, 11, "ИНН");
            Write(1, 12, "СНИЛС");
            Write(1, 13, "Клиент?");
            Write(1, 14, "Дата перехода в контрагенты");
            Write(1, 15, "Дата перехода в клиенты");
            Write(1, 16, "Тип документа(цифр.)");
            Write(1, 17, "Тип документа(строка)");
            Write(1, 18, "Серия");
            Write(1, 19, "Номер");
            Write(1, 20, "Документ выдан");
            Write(1, 21, "Код подразделения");
            Write(1, 22, "Дата выдачи ДУЛ");
            Write(1, 23, "Дата окончания ДУЛ");
            Write(1, 24, "Ответственное подразделение");
            Write(1, 25, "Гражданство");
            Write(1, 26, "Рабочий телефон");
            Write(1, 27, "Мобильный телефон");
            Write(1, 28, "Email");
            Write(1, 29, "Резерв1");
            Write(1, 30, "Резерв2");
            Write(1, 31, "Наличие выгодоприобретателя");
            Write(1, 32, "Принадлежит к ПДЛ");
            Write(1, 33, "Наличие бенефициарного владельца");
            Write(1, 34, "Деловая Репутация");
            Write(1, 35, "Финансовое положение");
            Write(1, 36, "Дата актуализации");
            Write(1, 37, "Сотрудник, проводивший актуализацию");
            Write(1, 38, "Тип обновления");
            Write(1, 39, "Дата последнего изменения");
            Write(1, 40, "Сотрудник, производивший изменение");
            Write(1, 41, "Пересекал границу");
            Write(1, 42, "Дата пересечения границы");
            Write(1, 43, "Код типа документа, разрешающего пребывание");
            Write(1, 44, "Расшифровка документа,");
            Write(1, 45, "Серия и номер");
            Write(1, 46, "Дата начала");
            Write(1, 47, "Дата окончания");
            Write(1, 48, "Тип МК");
            Write(1, 49, "Номер МК");
            Write(1, 50, "Дата начала МК");
            Write(1, 51, "Дата окончания МК");
            Write(1, 52, "Код уровня риска");
            Write(1, 53, "Расшифровка уровня риска");
            Write(1, 54, "Номер счета");
            Write(1, 55, "Тип счета");
            Write(1, 56, "Дата открытия счета");
            Write(1, 57, "Отделение счета");
            Write(1, 58, "Расширенное отделение счета");
            Write(1, 59, "Сотрудник, открывший счет");
            Write(1, 60, "Бух.режим");
            Write(1, 61, "Дата закрытия счета");
            Write(1, 62, "Адрес Prime");
            Write(1, 63, "Телефон Prime");
            Write(1, 64, "Адрес регистрации J");
            Write(1, 65, "Дата обновления адреса J");
            Write(1, 66, "Телефон J");
            Write(1, 67, "Адрес F");
            Write(1, 68, "Дата обновления адреса F");
            Write(1, 69, "Телефон F");
            Write(1, 70, "Статус записи");
            Write(1, 71, "Место рождения");
            Write(1, 72, "Код страны рождения");
            Write(1, 73, "Страна рождения");
            Write(1, 74, "Код региона рождения");
            Write(1, 75, "Тип региона рождения");
            Write(1, 76, "Регион рождения");
            Write(1, 77, "Тип города рождения");
            Write(1, 78, "Город рождения");
            Write(1, 79, "Тип населенного пункта рождения");
            Write(1, 80, "Населенный пункт рождения");
            Write(1, 81, "Место рождения полное");

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
