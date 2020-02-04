using System;
using System.Drawing;
using LinqToExcel;
using OfficeOpenXml;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Reflection;

namespace OFKO_Robot.Model
{
    /// <summary>
    /// Объект реализующий взаимодействие с Excel
    /// </summary>
    class ExcelWorker : Interfaces.IExcelWorker, IDisposable
    {
        ExcelPackage excel;
        ExcelWorksheet sheet;
        private readonly Equation equation;

        public ExcelWorker(Equation equation)
        {
            this.equation = equation;
        }

        /// <summary>
        /// Создать новый файл Excel
        /// </summary>
        /// <param name="filePath">Полное имя нового файла</param>
        public void CreateFile(string filePath)
        {
            excel = new ExcelPackage(new FileInfo(filePath));
            sheet = excel.Workbook.Worksheets.Add("Sheet1");
        }

        /// <summary>
        /// Очистить память от excel
        /// </summary>
        public void Dispose()
        {
            if (IsOn())
            {
                sheet.Dispose();
                excel.Dispose();
            }
        }

        /// <summary>
        /// Закрасить ячейку цветом
        /// </summary>
        /// <param name="x">Координат X</param>
        /// <param name="y">Координат Y</param>
        /// <param name="color">Выбранный цвет</param>
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
        /// <param name="filePath">Полное имя файла</param>
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

            }//xlsx разбираем Epplus
            else if (file.Extension.Equals(".xlsx"))
            {
                excel = new ExcelPackage(file);
                sheet = excel.Workbook.Worksheets[1];
            }
        }

        /// <summary>
        /// Данный метод реализует основной функционал программы.
        /// производится чтение исходного файла, 
        /// обращение к Equatoin
        /// запись результата в новый файл
        /// </summary>
        public void Work()
        {
            #region подготовка к работе
            placeHeaders();
            List<FL_data> mans = new List<FL_data>();
            Console.WriteLine("Reading data from Excel (Persons)");
            double onePercent = 100.00 / sheet.Dimension.Rows;
            double progress = 0;
            Statistic statistic = new Statistic(3);
            #endregion

            #region Поиск информации по каждому ПИНу в Equation
            for (int i = 2; i <= sheet.Dimension.Rows; i++)
            {
                Application.DoEvents();
                Console.WriteLine($"Done: {progress:F2}% ({i - 2}/{sheet.Dimension.Rows - 2})");
                mans.Add(equation.ReadData(sheet.Cells[i, 1].Text));
                progress = onePercent * i;
                LogWorker.ClearPrevConsoleLine();
                statistic.OperationDone(); //Сбор статистики об использовании робота.
            }
            #endregion

            Console.WriteLine("Done.");
            Console.WriteLine("Getting AccountsData");

            //Заполнение данных о счетах в получейся коллекции
            equation.fillAccounts(mans);

            Console.WriteLine("Done.");
            Console.WriteLine("Exporting data to .xlsx");

            #region Запись данных в Excel
            int rowIndex = 2;
            int column = 1;
            for (int i = 0; i < mans.Count; i++)
            {
                for (int j = 0; j < mans[i].Accounts.Count; j++)
                {
                    column = 1;
                    Write(rowIndex, column++, mans[i].Pin);
                    Write(rowIndex, column++, mans[i].Type);
                    Write(rowIndex, column++, mans[i].Identification);
                    Write(rowIndex, column++, mans[i].ApproveType);
                    Write(rowIndex, column++, mans[i].FIO);
                    Write(rowIndex, column++, NormalizeData(mans[i].BirthDate));
                    Write(rowIndex, column++, mans[i].BirthPlace);
                    Write(rowIndex, column++, mans[i].mnemonicFull);
                    Write(rowIndex, column++, mans[i].mnemonicShort);
                    Write(rowIndex, column++, mans[i].regCountry);
                    Write(rowIndex, column++, mans[i].Inn);
                    Write(rowIndex, column++, mans[i].SNILS);
                    Write(rowIndex, column++, mans[i].IsClient);
                    Write(rowIndex, column++, NormalizeData(mans[i].BecomeContragentDate));
                    Write(rowIndex, column++, NormalizeData(mans[i].BecomeClientDate));
                    Write(rowIndex, column++, mans[i].DocTypeInt);
                    Write(rowIndex, column++, mans[i].DocTypeStr);
                    Write(rowIndex, column++, mans[i].DocSeries);
                    Write(rowIndex, column++, mans[i].DocNumber);
                    Write(rowIndex, column++, mans[i].DocGiver);
                    Write(rowIndex, column++, mans[i].DocKPPCode);
                    Write(rowIndex, column++, NormalizeData(mans[i].DocOpenDate));
                    Write(rowIndex, column++, NormalizeData(mans[i].DocClosedDate));
                    Write(rowIndex, column++, mans[i].DocResponsibleUnit);
                    Write(rowIndex, column++, mans[i].Citizenship);
                    Write(rowIndex, column++, mans[i].WorkPhone);
                    Write(rowIndex, column++, mans[i].MobilePhone);
                    Write(rowIndex, column++, mans[i].Email);
                    Write(rowIndex, column++, mans[i].Reserve1);
                    Write(rowIndex, column++, mans[i].Reserve2);
                    Write(rowIndex, column++, mans[i].AnyBeneficiary);
                    Write(rowIndex, column++, mans[i].PDL);
                    Write(rowIndex, column++, mans[i].BeneficiaryOwner);
                    Write(rowIndex, column++, mans[i].Reputation);
                    Write(rowIndex, column++, mans[i].FinancialPosition);
                    Write(rowIndex, column++, NormalizeData(mans[i].ActualizationDate));
                    Write(rowIndex, column++, mans[i].ActualizationEmployee);
                    Write(rowIndex, column++, mans[i].UpdateType);
                    Write(rowIndex, column++, NormalizeData(mans[i].LastUpdatedDate));
                    Write(rowIndex, column++, mans[i].UpdateEmployee);
                    Write(rowIndex, column++, mans[i].IsCrossedTheBorder);
                    Write(rowIndex, column++, NormalizeData(mans[i].CrossBorderDate));
                    Write(rowIndex, column++, mans[i].ResidenceDocTypeInt);
                    Write(rowIndex, column++, mans[i].ResidenceDocTypeStr);
                    Write(rowIndex, column++, mans[i].ResidenceSerieAndNumber);
                    Write(rowIndex, column++, NormalizeData(mans[i].ResidenceDateStart));
                    Write(rowIndex, column++, NormalizeData(mans[i].ResidenceDateEnd));
                    Write(rowIndex, column++, mans[i].MigrationCardTypeInt);
                    Write(rowIndex, column++, mans[i].MigrationCardNumber);
                    Write(rowIndex, column++, NormalizeData(mans[i].MigrationCardDateStart));
                    Write(rowIndex, column++, NormalizeData(mans[i].MigrationCardDateEnd));
                    Write(rowIndex, column++, mans[i].RiskLevelInt.ToString());
                    Write(rowIndex, column++, mans[i].RiskLevelStr);
                    Write(rowIndex, column++, mans[i].Accounts[j].Number);
                    Write(rowIndex, column++, mans[i].Accounts[j].Type);
                    Write(rowIndex, column++, mans[i].Accounts[j].OpenDate.ToString("dd.MM.yyyy"));
                    Write(rowIndex, column++, mans[i].Accounts[j].Mnemonic);
                    Write(rowIndex, column++, mans[i].Accounts[j].MnemonicFull);
                    Write(rowIndex, column++, mans[i].Accounts[j].EmployeeOpenedFIO);
                    Write(rowIndex, column++, mans[i].Accounts[j].AccountingMode);
                    Write(rowIndex, column++, NormalizeData(mans[i].Accounts[j].CloseDate));
                    Write(rowIndex, column++, mans[i].RegistrationAddress);
                    Write(rowIndex, column++, mans[i].PrimePhone);
                    Write(rowIndex, column++, mans[i].RegistrationAddressFull);
                    Write(rowIndex, column++, NormalizeData(mans[i].RegistrationUpdateDate));
                    Write(rowIndex, column++, mans[i].Registration_Index);
                    Write(rowIndex, column++, mans[i].Registration_Country);
                    Write(rowIndex, column++, mans[i].Registration_RegionCode);
                    Write(rowIndex, column++, mans[i].Registration_RegionType);
                    Write(rowIndex, column++, mans[i].Registration_RegionName);
                    Write(rowIndex, column++, mans[i].Registration_Town);
                    Write(rowIndex, column++, mans[i].Registration_LocalityType);
                    Write(rowIndex, column++, mans[i].Registration_LocalityName);
                    Write(rowIndex, column++, mans[i].Registration_StreetType);
                    Write(rowIndex, column++, mans[i].Registration_StreetName);
                    Write(rowIndex, column++, mans[i].Registration_HouseNumber);
                    Write(rowIndex, column++, mans[i].Registration_BuildingNumber);
                    Write(rowIndex, column++, mans[i].Registration_FlatNumber);
                    Write(rowIndex, column++, mans[i].RegistrationPhone);
                    Write(rowIndex, column++, mans[i].FactAddress);
                    Write(rowIndex, column++, mans[i].Fact_Index);
                    Write(rowIndex, column++, mans[i].Fact_Country);
                    Write(rowIndex, column++, mans[i].Fact_RegionCode);
                    Write(rowIndex, column++, mans[i].Fact_RegionType);
                    Write(rowIndex, column++, mans[i].Fact_RegionName);
                    Write(rowIndex, column++, mans[i].Fact_Town);
                    Write(rowIndex, column++, mans[i].Fact_LocalityType);
                    Write(rowIndex, column++, mans[i].Fact_LocalityName);
                    Write(rowIndex, column++, mans[i].Fact_StreetType);
                    Write(rowIndex, column++, mans[i].Fact_StreetName);
                    Write(rowIndex, column++, mans[i].Fact_HouseNumber);
                    Write(rowIndex, column++, mans[i].Fact_BuildingNumber);
                    Write(rowIndex, column++, mans[i].Fact_FlatNumber);
                    Write(rowIndex, column++, NormalizeData(mans[i].FactAddressUpdateDate));
                    Write(rowIndex, column++, mans[i].FactAddressPhone);
                    Write(rowIndex, column++, mans[i].RecordStatus);
                    Write(rowIndex, column++, mans[i].BirthPlace);
                    Write(rowIndex, column++, mans[i].BirthPlaceCountryCode);
                    Write(rowIndex, column++, mans[i].BirthPlaceCountry);
                    Write(rowIndex, column++, mans[i].BirthPlaceRegionCode);
                    Write(rowIndex, column++, mans[i].BirthPlaceRegionType);
                    Write(rowIndex, column++, mans[i].BirthPlaceRegion);
                    Write(rowIndex, column++, mans[i].BirthPlaceCityType);
                    Write(rowIndex, column++, mans[i].BirthPlaceCity);
                    Write(rowIndex, column++, mans[i].BirthPlaceLocalityType);
                    Write(rowIndex, column++, mans[i].BirthPlaceLocality);
                    Write(rowIndex, column++, mans[i].BirthPlaceFull);

                    rowIndex++;
                }
            }
            #endregion

            //Сохранение файла Excel
            Save();
            equation.CloseConnection();
            //отправка статистики на сервер
            statistic.Commit();
        }

        /// <summary>
        /// Метод отсекающий некорректные значения даты.
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns>возвращает пустую строку, если дата была не корректна, и дату в 
        /// формате ДД.ММ.ГГГГ в случае корректной даты</returns>
        private string NormalizeData(DateTime dateTime)
        {
            return dateTime.Year < 1000 ? string.Empty : dateTime.ToString("dd.MM.yyyy");
        }

        /// <summary>
        /// Разместить заголовки столбцов
        /// </summary>
        private void placeHeaders()
        {
            int row = 1;
            int column = 1;
            Write(row, column++, "Пин");
            Write(row, column++, "Тип");
            Write(row, column++, "Идентификация");
            Write(row, column++, "Тип подтверждения");
            Write(row, column++, "ФИО");
            Write(row, column++, "Дата рождения");
            Write(row, column++, "Место рождения");
            Write(row, column++, "Мнемоника_расшифровка");
            Write(row, column++, "Мнемоника");
            Write(row, column++, "Страна регистрации");
            Write(row, column++, "ИНН");
            Write(row, column++, "СНИЛС");
            Write(row, column++, "Клиент?");
            Write(row, column++, "Дата перехода в контрагенты");
            Write(row, column++, "Дата перехода в клиенты");
            Write(row, column++, "Тип документа(цифр.)");
            Write(row, column++, "Тип документа(строка)");
            Write(row, column++, "Серия");
            Write(row, column++, "Номер");
            Write(row, column++, "Документ выдан");
            Write(row, column++, "Код подразделения");
            Write(row, column++, "Дата выдачи ДУЛ");
            Write(row, column++, "Дата окончания ДУЛ");
            Write(row, column++, "Ответственное подразделение");
            Write(row, column++, "Гражданство");
            Write(row, column++, "Рабочий телефон");
            Write(row, column++, "Мобильный телефон");
            Write(row, column++, "Email");
            Write(row, column++, "Резерв1");
            Write(row, column++, "Резерв2");
            Write(row, column++, "Наличие выгодоприобретателя");
            Write(row, column++, "Принадлежит к ПДЛ");
            Write(row, column++, "Наличие бенефициарного владельца");
            Write(row, column++, "Деловая Репутация");
            Write(row, column++, "Финансовое положение");
            Write(row, column++, "Дата актуализации");
            Write(row, column++, "Сотрудник, проводивший актуализацию");
            Write(row, column++, "Тип обновления");
            Write(row, column++, "Дата последнего изменения");
            Write(row, column++, "Сотрудник, производивший изменение");
            Write(row, column++, "Пересекал границу");
            Write(row, column++, "Дата пересечения границы");
            Write(row, column++, "Код типа документа, разрешающего пребывание");
            Write(row, column++, "Расшифровка документа,");
            Write(row, column++, "Серия и номер");
            Write(row, column++, "Дата начала");
            Write(row, column++, "Дата окончания");
            Write(row, column++, "Тип МК");
            Write(row, column++, "Номер МК");
            Write(row, column++, "Дата начала МК");
            Write(row, column++, "Дата окончания МК");
            Write(row, column++, "Код уровня риска");
            Write(row, column++, "Расшифровка уровня риска");
            Write(row, column++, "Номер счета");
            Write(row, column++, "Тип счета");
            Write(row, column++, "Дата открытия счета");
            Write(row, column++, "Отделение счета");
            Write(row, column++, "Расширенное отделение счета");
            Write(row, column++, "Сотрудник, открывший счет");
            Write(row, column++, "Бух.режим");
            Write(row, column++, "Дата закрытия счета");
            Write(row, column++, "Адрес Prime");
            Write(row, column++, "Телефон Prime");
            Write(row, column++, "Адрес регистрации J");
            Write(row, column++, "Дата обновления адреса J");
            Write(row, column++, "Индекс J");
            Write(row, column++, "Страна J");
            Write(row, column++, "Код региона J");
            Write(row, column++, "Тип региона J");
            Write(row, column++, "Наименование района J");
            Write(row, column++, "Город J");
            Write(row, column++, "Тип населенного пункта J");
            Write(row, column++, "Населенный пункт J");
            Write(row, column++, "Тип улицы J");
            Write(row, column++, "Наименование улицы J");
            Write(row, column++, "Номер дома J");
            Write(row, column++, "Корпус/строение J");
            Write(row, column++, "Квартира/комната J");
            Write(row, column++, "Телефон J");
            Write(row, column++, "Адрес F");
            Write(row, column++, "Индекс F");
            Write(row, column++, "Страна F");
            Write(row, column++, "Код региона F");
            Write(row, column++, "Тип региона F");
            Write(row, column++, "Наименование района F");
            Write(row, column++, "Город F");
            Write(row, column++, "Тип населенного пункта F");
            Write(row, column++, "Населенный пункт F");
            Write(row, column++, "Тип улицы F");
            Write(row, column++, "Наименование улицы F");
            Write(row, column++, "Номер дома F");
            Write(row, column++, "Корпус/строение F");
            Write(row, column++, "Квартира/комната F");
            Write(row, column++, "Дата обновления адреса F");
            Write(row, column++, "Телефон F");
            Write(row, column++, "Статус записи");
            Write(row, column++, "Место рождения");
            Write(row, column++, "Код страны рождения");
            Write(row, column++, "Страна рождения");
            Write(row, column++, "Код региона рождения");
            Write(row, column++, "Тип региона рождения");
            Write(row, column++, "Регион рождения");
            Write(row, column++, "Тип города рождения");
            Write(row, column++, "Город рождения");
            Write(row, column++, "Тип населенного пункта рождения");
            Write(row, column++, "Населенный пункт рождения");
            Write(row, column++, "Место рождения полное");

        }

        /// <summary>
        /// Получить значение ячейки
        /// </summary>
        /// <param name="x">Координат X</param>
        /// <param name="y">Координат Y</param>
        /// <returns></returns>
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

        /// <summary>
        /// Сохранить текущий файл Excel
        /// </summary>
        /// <param name="filePath">Полное имя файла</param>
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

        /// <summary>
        /// Сохранить текущие изменения в файле excel
        /// </summary>
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

        /// <summary>
        /// Записать значение в ячейку Excel
        /// </summary>
        /// <param name="x">Координат X</param>
        /// <param name="y">Координат Y</param>
        /// <param name="text">Новое значение</param>
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

        /// <summary>
        /// Проверяет наличие объекта excel и sheet
        /// </summary>
        /// <returns>True если существует excel и sheet
        /// False если один из объектов является null</returns>
        private bool IsOn()
        {
            if (excel == null || sheet == null)
                return false;
            return true;
        }
    }
}
