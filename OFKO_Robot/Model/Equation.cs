using EquationLibrary;
using System;
using System.Collections.Generic;
using System.Text;

namespace OFKO_Robot.Model
{
    /// <summary>
    /// Класс описывает логику взаимодействия с Equation
    /// </summary>
    class Equation
    {
        // Наличие соединения с PCCOMM
        public bool connected = true;

        //Имя текущей сессии
        char sessionName = 'A';

        public Equation()
        {
            #region getConnection
            int attempCount = 0;
            Console.WriteLine("Connecting to Equation");
            Console.WriteLine($"Trying connect to {sessionName}...");
            while (EUCL.Connect(sessionName.ToString()) != 0)
            {
                if (attempCount++ > 5)
                {
                    connected = false;
                    Console.WriteLine("Connection failed");
                    break;
                }
                sessionName++;
                Console.WriteLine($"Trying connect to {sessionName}...");
            }
            if (connected)
            {
                ClearScreen();
                Console.WriteLine("Connected.");
            }
            #endregion
        }

        /// <summary>
        /// Привести экран в начальное положение (закрыть все опции ит.п.)
        /// </summary>
        public void ClearScreen() => EUCL.ClearScreen();

        /// <summary>
        /// Получает информацию по конкретному физ.лицу
        /// </summary>
        /// <param name="pin">ПИН клиента</param>
        /// <returns>Возвращает объект FL_data</returns>
        public FL_data ReadData(string pin)
        {
            FL_data fl = new FL_data();
            fl.Pin = pin;

            #region ПФ1_function

            send(21, 17, "ПФ1");
            pEnter();
            send(3, 34, fl.Pin);
            pEnter();
            fl.Type = EUCL.ReadScreen(4, 34, 2);
            fl.Identification = ClearWhiteSpacesStr(EUCL.ReadScreen(2, 14, 30));// на чтении экрана в EQ(что бы не пришлось читать экран дважды)
            fl.ApproveType = ClearWhiteSpacesStr(EUCL.ReadScreen(2, 60, 30));
            fl.FIO = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(7, 34, 35))} {ClearWhiteSpacesStr(EUCL.ReadScreen(8, 34, 35))} {ClearWhiteSpacesStr(EUCL.ReadScreen(9, 34, 35))}";
            fl.BirthDate = DateTime.Parse(EUCL.ReadScreen(10, 34, 11));
            fl.mnemonicFull = ClearWhiteSpacesStr(EUCL.ReadScreen(15, 41, 35));
            fl.mnemonicShort = EUCL.ReadScreen(15, 34, 4);
            fl.regCountry = EUCL.ReadScreen(16, 34, 2);
            fl.Inn = EUCL.ReadScreen(18, 34, 12);
            fl.SNILS = EUCL.ReadScreen(18, 57, 18).Replace(" - ", "-");
            if (fl.SNILS.Equals("   -   -      ")) { fl.SNILS = string.Empty; }
            fl.IsClient = EUCL.ReadScreen(21, 67, 1);
            DateTime.TryParse(EUCL.ReadScreen(20, 34, 11), out fl.BecomeContragentDate);
            DateTime.TryParse(EUCL.ReadScreen(21, 34, 11), out fl.BecomeClientDate);
            pEnter();
            fl.DocTypeInt = EUCL.ReadScreen(5, 34, 3);
            fl.DocTypeStr = ClearWhiteSpacesStr(EUCL.ReadScreen(5, 41, 35));
            fl.DocSeries = ClearWhiteSpacesStr(EUCL.ReadScreen(7, 34, 10));
            fl.DocNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 34, 35));
            DateTime.TryParse(EUCL.ReadScreen(6, 34, 11), out fl.DocOpenDate);
            DateTime.TryParse(EUCL.ReadScreen(6, 67, 11), out fl.DocClosedDate);
            fl.DocGiver = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(9, 34, 35))}{ClearWhiteSpacesStr(EUCL.ReadScreen(10, 2, 73))}{ClearWhiteSpacesStr(EUCL.ReadScreen(11, 2, 73))}{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 2, 73))}";
            fl.DocKPPCode = EUCL.ReadScreen(13, 34, 7);
            fl.DocResponsibleUnit = EUCL.ReadScreen(14, 34, 3);
            fl.DocResponsibleUnitFull = EUCL.ReadScreen(14, 41, 35);
            fl.Citizenship = EUCL.ReadScreen(16, 34, 2);

            pEnter();
            pEnter();
            fl.WorkPhone = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 34, 35));
            fl.MobilePhone = ClearWhiteSpacesStr(EUCL.ReadScreen(9, 34, 35));
            fl.Email = ClearWhiteSpacesStr(EUCL.ReadScreen(11, 34, 35));
            pEnter();
            fl.Reserve1 = ClearWhiteSpacesStr(EUCL.ReadScreen(19, 46, 35));
            fl.Reserve2 = ClearWhiteSpacesStr(EUCL.ReadScreen(20, 46, 35));
            pEnter();
            pEnter();
            fl.AnyBeneficiary = EUCL.ReadScreen(6, 37, 1);
            fl.PDL = EUCL.ReadScreen(7, 37, 1);
            fl.BeneficiaryOwner = EUCL.ReadScreen(8, 37, 1);
            fl.Reputation = ClearWhiteSpacesStr(EUCL.ReadScreen(14, 37, 35));
            fl.FinancialPosition = ClearWhiteSpacesStr(EUCL.ReadScreen(15, 37, 35));
            fl.ActualizationDate = parseCustom(EUCL.ReadScreen(18, 37, 6));
            fl.ActualizationEmployee = ClearWhiteSpacesStr(EUCL.ReadScreen(18, 45, 35));
            fl.UpdateType = ClearWhiteSpacesStr(EUCL.ReadScreen(20, 40, 35));
            fl.LastUpdatedDate = parseCustom(EUCL.ReadScreen(21, 37, 6));
            fl.UpdateEmployee = EUCL.ReadScreen(21, 73, 4);
            pEnter();
            pEnter();
            send(13, 77, "Y");
            pEnter();
            fl.IsCrossedTheBorder = EUCL.ReadScreen(9, 32, 1);
            DateTime.TryParse(EUCL.ReadScreen(9, 54, 11), out fl.CrossBorderDate);
            fl.ResidenceDocTypeInt = EUCL.ReadScreen(15, 5, 2);
            fl.ResidenceDocTypeStr = ClearWhiteSpacesStr(EUCL.ReadScreen(15, 10, 20));
            fl.ResidenceSerieAndNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(16, 10, 20));
            fl.ResidenceDateStart = parseCustom(EUCL.ReadScreen(15, 31, 6));
            fl.ResidenceDateEnd = parseCustom(EUCL.ReadScreen(16, 31, 6));
            fl.MigrationCardTypeInt = EUCL.ReadScreen(15, 41, 2);
            fl.MigrationCardNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(16, 46, 20));
            fl.MigrationCardDateStart = parseCustom(EUCL.ReadScreen(15, 67, 6));
            fl.MigrationCardDateEnd = parseCustom(EUCL.ReadScreen(16, 67, 6));
            pEnter();
            int.TryParse(EUCL.ReadScreen(5, 34, 1), out fl.RiskLevelInt);
            fl.RiskLevelStr = ClearWhiteSpacesStr(EUCL.ReadScreen(5, 36, 40));
            pEnter();
            if (EUCL.ReadScreen(1, 31, 24).Equals("Список повторяющихся ИНН"))
            { pEnter(); }
            EUCL.SendStr("@3");

            #endregion

            #region ПАД_function

            send(21, 17, "ПАД");
            pEnter();
            send(7, 29, "J");
            send(3, 29, fl.Pin);
            pEnter();

            StringBuilder builder = new StringBuilder();

            if (EUCL.ReadScreen(24, 3, 3).Equals("KSM")) //адрес не существует
            {
                fl.RegistrationAddressFull = "Отсутствует";
            }
            else
            {
                builder.Append($"гор.{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 29, 50))}, ");
                builder.Append($"{ClearWhiteSpacesStr(EUCL.ReadScreen(15, 33, 20))} ");
                builder.Append($"{ClearWhiteSpacesStr(EUCL.ReadScreen(16, 29, 30))}, ");
                builder.Append($"д. {ClearWhiteSpacesStr(EUCL.ReadScreen(17, 14, 6))}, ");
                builder.Append($"корп. {ClearWhiteSpacesStr(EUCL.ReadScreen(17, 47, 6))}, ");
                builder.Append($"кв. {ClearWhiteSpacesStr(EUCL.ReadScreen(17, 73, 6))}");
                fl.RegistrationAddressFull = builder.ToString();
                fl.Registration_Index = ClearWhiteSpacesStr(EUCL.ReadScreen(7,29,15));
                fl.Registration_Country = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 29, 35));
                fl.Registration_RegionCode = ClearWhiteSpacesStr(EUCL.ReadScreen(9, 29, 35));
                fl.Registration_RegionType = ClearWhiteSpacesStr(EUCL.ReadScreen(10,29,35));
                fl.Registration_RegionName = ClearWhiteSpacesStr(EUCL.ReadScreen(11,29,35));
                fl.Registration_Town = ClearWhiteSpacesStr(EUCL.ReadScreen(12, 29, 35));
                fl.Registration_LocalityType = ClearWhiteSpacesStr(EUCL.ReadScreen(13, 29, 35));
                fl.Registration_LocalityName = ClearWhiteSpacesStr(EUCL.ReadScreen(14, 29, 35));
                fl.Registration_StreetType = ClearWhiteSpacesStr(EUCL.ReadScreen(15, 29, 35));
                fl.Registration_StreetName = ClearWhiteSpacesStr(EUCL.ReadScreen(16, 29, 35));
                fl.Registration_HouseNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 14, 8));
                fl.Registration_BuildingNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 46, 8));
                fl.Registration_FlatNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 73, 8));
                builder.Clear();
                DateTime.TryParse(EUCL.ReadScreen(5, 37, 11), out fl.RegistrationUpdateDate);
                pEnter();
                fl.RegistrationPhone = ClearWhiteSpacesStr(EUCL.ReadScreen(6, 29, 15));
                EUCL.SendStr("@c");
                EUCL.Wait();
                EUCL.SendStr("@c");
                EUCL.Wait();
            }

            send(7, 29, "F");
            pEnter();

            if (EUCL.ReadScreen(24, 3, 3).Equals("KSM")) //адрес не существует
            {
                fl.FactAddress = "Отсутствует";
            }

            else
            {
                builder.Append($"гор.{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 29, 50))}, ");
                builder.Append($"{ClearWhiteSpacesStr(EUCL.ReadScreen(15, 33, 20))} ");
                builder.Append($"{ClearWhiteSpacesStr(EUCL.ReadScreen(16, 29, 30))}, ");
                builder.Append($"д. {ClearWhiteSpacesStr(EUCL.ReadScreen(17, 14, 6))}, ");
                builder.Append($"корп. {ClearWhiteSpacesStr(EUCL.ReadScreen(17, 47, 6))}, ");
                builder.Append($"кв. {ClearWhiteSpacesStr(EUCL.ReadScreen(17, 73, 6))}");
                fl.FactAddress = builder.ToString();
                fl.Fact_Index = ClearWhiteSpacesStr(EUCL.ReadScreen(7, 29, 15));
                fl.Fact_Country = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 29, 35));
                fl.Fact_RegionCode = ClearWhiteSpacesStr(EUCL.ReadScreen(9, 29, 35));
                fl.Fact_RegionType = ClearWhiteSpacesStr(EUCL.ReadScreen(10, 29, 35));
                fl.Fact_RegionName = ClearWhiteSpacesStr(EUCL.ReadScreen(11, 29, 35));
                fl.Fact_Town = ClearWhiteSpacesStr(EUCL.ReadScreen(12, 29, 35));
                fl.Fact_LocalityType = ClearWhiteSpacesStr(EUCL.ReadScreen(13, 29, 35));
                fl.Fact_LocalityName = ClearWhiteSpacesStr(EUCL.ReadScreen(14, 29, 35));
                fl.Fact_StreetType = ClearWhiteSpacesStr(EUCL.ReadScreen(15, 29, 35));
                fl.Fact_StreetName = ClearWhiteSpacesStr(EUCL.ReadScreen(16, 29, 35));
                fl.Fact_HouseNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 14, 8));
                fl.Fact_BuildingNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 46, 8));
                fl.Fact_FlatNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 73, 8));
                DateTime.TryParse(EUCL.ReadScreen(5, 37, 11), out fl.FactAddressUpdateDate);
                pEnter();
                fl.FactAddressPhone = ClearWhiteSpacesStr(EUCL.ReadScreen(6, 29, 15));
                EUCL.SendStr("@c");
                EUCL.Wait();
                EUCL.SendStr("@c");
                EUCL.Wait();
            }

            send(7, 29, " ");
            pEnter();

            if (!EUCL.ReadScreen(24, 3, 3).Equals("KSM")) //адрес не существует
            {
                fl.RegistrationAddress = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(7, 29, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(8, 29, 40))}";
                fl.PrimePhone = ClearWhiteSpacesStr(EUCL.ReadScreen(14, 29, 35));
            }

            EUCL.SendStr("@c");
            EUCL.Wait();

            #endregion

            #region MR9_function

            send(21, 17, "MR9");
            pEnter();
            send(3, 17, fl.Pin);
            pEnter();
            if (string.IsNullOrWhiteSpace(EUCL.ReadScreen(4, 4, 11)))
            {
                fl.RecordStatus = string.Empty;
                fl.BirthPlace = string.Empty;
                fl.BirthPlaceCountryCode = string.Empty;
                fl.BirthPlaceCountry = string.Empty;
                fl.BirthPlaceRegionCode = string.Empty;
                fl.BirthPlaceRegionType = string.Empty;
                fl.BirthPlaceRegion = string.Empty;
                fl.BirthPlaceCityType = string.Empty;
                fl.BirthPlaceCity = string.Empty;
                fl.BirthPlaceLocalityType = string.Empty;
                fl.BirthPlaceLocality = string.Empty;
                fl.BirthPlaceFull = string.Empty;
            }
            else
            {
                fl.RecordStatus = ClearWhiteSpacesStr(EUCL.ReadScreen(4, 4, 11));
                send(4, 2, "1");
                EUCL.Wait();
                fl.BirthPlace = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(6, 17, 50))}{ClearWhiteSpacesStr(EUCL.ReadScreen(7, 17, 50))}{ClearWhiteSpacesStr(EUCL.ReadScreen(8, 17, 50))}";
                pEnter();
                fl.BirthPlaceCountryCode = EUCL.ReadScreen(10, 18, 3);
                fl.BirthPlaceCountry = ClearWhiteSpacesStr(EUCL.ReadScreen(12, 17, 40));
                fl.BirthPlaceRegionCode = ClearWhiteSpacesStr(EUCL.ReadScreen(17, 17, 40));
                pEnter();
                fl.BirthPlaceRegionType = ClearWhiteSpacesStr(EUCL.ReadScreen(10, 27, 30));
                fl.BirthPlaceRegion = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(13, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(14, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(15, 17, 40))}";
                pEnter();
                fl.BirthPlaceCityType = ClearWhiteSpacesStr(EUCL.ReadScreen(10, 27, 30));
                fl.BirthPlaceCity = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(13, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(14, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(15, 17, 40))}";
                pEnter();
                fl.BirthPlaceLocalityType = ClearWhiteSpacesStr(EUCL.ReadScreen(10, 40, 0));
                fl.BirthPlaceLocality = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(13, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(14, 17, 40))}{ClearWhiteSpacesStr(EUCL.ReadScreen(15, 17, 40))}";
                pEnter();
                fl.BirthPlaceFull = ClearWhiteSpacesStr(EUCL.ReadScreen(5, 32, 40));
            }
            ClearScreen();
            #endregion

            return fl;
        }

        /// <summary>
        /// Получить информацию о всех счетах с бух.режимом 100
        /// </summary>
        /// <param name="data">Коллекция физ.лиц</param>
        public void fillAccounts(List<FL_data> data)
        {
            EUCL.ClearScreen();
            send(21, 17, "ppp");
            pEnter();
            pEnter();
            send(6, 2, "1");
            EUCL.Wait();
            send(5, 5, "40*  ");
            send(5, 77, "*   ");

            double onePercent = 100.00 / data.Count;
            double currentProgress = 0.0;
            int counter = 0;
            foreach (FL_data item in data)
            {
                currentProgress += onePercent;
                Console.WriteLine($"Done: {currentProgress:F2}% ({counter++}/{data.Count})");
                send(7, 69, item.Pin);
                send(5, 71, "100");
                pEnter();
                List<Account> accList = new List<Account>();

                for (int i = 8; i < 20; i += 2)
                {
                    Account acc = new Account();
                    acc.Number = EUCL.ReadScreen(i, 5, 24).Replace(".", string.Empty);
                    acc.OpenDate = parseCustom(EUCL.ReadScreen(i, 56, 6));
                    if (string.IsNullOrWhiteSpace(acc.Number))
                    {
                        break;
                    }
                    send(i, 2, "1");
                    EUCL.Wait();
                    acc.Type = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 34, 35));
                    acc.Mnemonic = EUCL.ReadScreen(10, 34, 4);
                    acc.MnemonicFull = ClearWhiteSpacesStr(EUCL.ReadScreen(10, 40, 35));
                    EUCL.SendStr("@l"); //F21
                    EUCL.Wait();
                    acc.AccountingMode = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 18, 35));
                    DateTime.TryParse(EUCL.ReadScreen(4, 70, 11), out acc.OpenDate);
                    DateTime.TryParse(EUCL.ReadScreen(5, 70, 11), out acc.CloseDate);
                    if (string.IsNullOrWhiteSpace(ClearWhiteSpacesStr(EUCL.ReadScreen(16, 18, 35))))
                    {
                        acc.EmployeeOpenedFIO = EUCL.ReadScreen(9, 18, 4);
                    }
                    else
                    {
                        send(16, 2, "1");
                        EUCL.Wait();
                        acc.EmployeeOpenedFIO = EUCL.ReadScreen(6, 26, 4);
                        EUCL.SendStr("@c"); //F12
                        EUCL.Wait();
                    }
                    EUCL.SendStr("@3"); //F3
                    EUCL.Wait();
                    EUCL.SendStr("@c"); //F12
                    EUCL.Wait();
                    accList.Add(acc);
                }
                LogWorker.ClearPrevConsoleLine();
                item.Accounts = accList;
            }
        }

        /// <summary>
        /// Приводит строку формата 311220 в формат DateTime
        /// </summary>
        /// <param name="dateString">текстовое значение даты</param>
        /// <returns>Возвращает объект DateTime</returns>
        private DateTime parseCustom(string dateString)
        {
            try
            {
                string tempStr = $"{dateString.Substring(0, 2)}.{dateString.Substring(2, 2)}.{dateString.Substring(4, 2)}";
                return DateTime.Parse(tempStr);
            }
            catch
            {
                return new DateTime();
            }
        }

        /// <summary>
        /// Отправить нажатие Enter и ожидать готовности ввода от PCCOMM
        /// </summary>
        private void pEnter()
        {
            EUCL.SendStr("@E");
            EUCL.Wait();
        }

        /// <summary>
        /// Отправить строку или команду по указанным координатам
        /// </summary>
        /// <param name="x">координат X</param>
        /// <param name="y">координат Y</param>
        /// <param name="Text">Отправляемая строка</param>
        /// <param name="size">Размер экрана</param>
        private void send(int x, int y, string Text, int size = 80)
        {
            EUCL.SetCursorPos(x, y, size);
            EUCL.SendStr(Text);
        }

        /// <summary>
        /// Очищает строку от двойных пробелов
        /// </summary>
        /// <param name="text">Редактируемая строка</param>
        /// <returns>Возвращает строку с удаленными двойными пробелами</returns>
        private string ClearWhiteSpacesStr(string text)
        {
            return text.Replace("  ", "");
        }

        /// <summary>
        /// Закрыть соединение с PCOMM
        /// </summary>
        public void CloseConnection()
        {
            EUCL.ClearScreen();
            EUCL.Disconnect(sessionName.ToString());
            Console.WriteLine("Соединение с Equation закрыто");
        }
    }
}
