using EquationLibrary;
using System;
using System.Collections.Generic;
using System.Text;

namespace OFKO_Robot.Model
{
    class Equation
    {
        public bool connected = true;
        char sessionName = 'A';
        public Equation()
        {
            #region getConnection
            int attempCount = 0;
            Console.WriteLine("Connecting to Equation");
            Console.WriteLine($"Trying connect to {sessionName}...");
            while (EUCL.Connect(sessionName.ToString())!=0)
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
        public void ClearScreen() => EUCL.ClearScreen();
        public FL_data ReadData(string pin)
        {
            FL_data fl = new FL_data();
            fl.Pin = pin;
            send(21, 17, "ПФ1");
            pEnter();
            send(3, 34, fl.Pin);
            pEnter();
            fl.Type = EUCL.ReadScreen(4, 34, 2);
            fl.Identification = ClearWhiteSpacesStr(EUCL.ReadScreen(2, 14, 30));// на чтении экрана в EQ(что бы не пришлось читать экран дважды)
            fl.ApproveType = ClearWhiteSpacesStr(EUCL.ReadScreen(2, 60, 30));
            fl.FIO = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(7,34, 35))} {ClearWhiteSpacesStr(EUCL.ReadScreen(8, 34, 35))} {ClearWhiteSpacesStr(EUCL.ReadScreen(9, 34, 35))}";
            fl.BirthDate = DateTime.Parse(EUCL.ReadScreen(10, 34, 11));
            fl.mnemonicFull = ClearWhiteSpacesStr(EUCL.ReadScreen(15, 41, 35));
            fl.mnemonicShort = EUCL.ReadScreen(15, 34, 4);
            fl.regCountry = EUCL.ReadScreen(16, 34, 2);
            fl.Inn = EUCL.ReadScreen(18, 34, 12);
            fl.SNILS = EUCL.ReadScreen(18, 57, 18).Replace(" - ", "-");
            if(fl.SNILS.Equals("   -   -      ")) { fl.SNILS = string.Empty; }
            fl.IsClient = EUCL.ReadScreen(21, 67, 1).Equals("Y") ? true : false;
            DateTime.TryParse(EUCL.ReadScreen(20, 34, 11), out fl.BecomeContragentDate);
            DateTime.TryParse(EUCL.ReadScreen(21, 34, 11), out fl.BecomeClientDate);
            pEnter();
            fl.DocTypeInt = EUCL.ReadScreen(5, 34, 3);
            fl.DocTypeStr = ClearWhiteSpacesStr(EUCL.ReadScreen(5, 41, 35));
            fl.DocSeries = ClearWhiteSpacesStr(EUCL.ReadScreen(7, 34, 10));
            fl.DocNumber = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 34, 35));
            fl.DocGiver = $"{ClearWhiteSpacesStr(EUCL.ReadScreen(9,34,35))}{ClearWhiteSpacesStr(EUCL.ReadScreen(10, 2, 73))}{ClearWhiteSpacesStr(EUCL.ReadScreen(11, 2, 73))}{ClearWhiteSpacesStr(EUCL.ReadScreen(12, 2, 73))}";
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
            fl.AnyBeneficiary = EUCL.ReadScreen(6, 37, 1).Equals("Y") ? true : false;
            fl.PDL = EUCL.ReadScreen(7, 37, 1).Equals("Y") ? true : false;
            fl.BeneficiaryOwner = EUCL.ReadScreen(8, 37, 1).Equals("Y") ? true : false;
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
            fl.IsCrossedTheBorder = EUCL.ReadScreen(9, 32, 1).Equals("Y") ? true : false;
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
            if (EUCL.ReadScreen(1,31,24).Equals("Список повторяющихся ИНН"))
            { pEnter(); }
            EUCL.SendStr("@3");
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
                builder.Clear();
                DateTime.TryParse(EUCL.ReadScreen(5, 37, 11), out fl.RegistrationUpdateDate);
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
                DateTime.TryParse(EUCL.ReadScreen(5, 37, 11), out fl.FactAddressUpdateDate);
                EUCL.SendStr("@c");
                EUCL.Wait();

            }
            send(7, 29, " ");
            pEnter();
            if (EUCL.ReadScreen(24, 3, 3).Equals("KSM")) //адрес не существует
            {
                fl.RegistrationAddress = "Отсутствует";
                fl.MobilePhone = "Отсутствует";
            }
            else
            {
                fl.RegistrationAddress = ClearWhiteSpacesStr(EUCL.ReadScreen(7, 29, 40));
                fl.MobilePhone = ClearWhiteSpacesStr(EUCL.ReadScreen(14, 29, 15));
            }
            EUCL.SendStr("@c");
            EUCL.Wait();

            /*MR9 - 
            RecordStatus;
            BirthPlace;
            BirthPlaceCountryCode;
            BirthPlaceCountry;
            BirthPlaceRegionCode;
            BirthPlaceRegionType;
            BirthPlaceRegion;
            BirthPlaceCityType;
            BirthPlaceCity;
            BirthPlaceLocalityType;
            BirthPlaceLocality;
            BirthPlaceFull;*/
            return fl;
        }
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
                
            foreach (FL_data item in data)
            {
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
                    acc.Type = EUCL.ReadScreen(8, 34, 2);
                    acc.Mnemonic = EUCL.ReadScreen(10, 34, 4);
                    acc.MnemonicFull = ClearWhiteSpacesStr(EUCL.ReadScreen(10, 40, 35));
                    acc.EmployeeOpenedFIO = EUCL.ReadScreen(16, 34, 4);
                    EUCL.SendStr("@l"); //F21
                    EUCL.Wait();
                    acc.AccountingMode = ClearWhiteSpacesStr(EUCL.ReadScreen(8, 18, 35));
                    DateTime.TryParse(EUCL.ReadScreen(4, 70, 11), out acc.OpenDate);
                    DateTime.TryParse(EUCL.ReadScreen(5, 70, 11), out acc.CloseDate);
                    EUCL.SendStr("@3"); //F3
                    EUCL.Wait();
                    EUCL.SendStr("@c"); //F12
                    EUCL.Wait();
                    accList.Add(acc);
                }

                item.Accounts = accList;
            }
        }
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
        private void pEnter()
        {
            EUCL.SendStr("@E");
            EUCL.Wait();
        }
        private void send(int x, int y, string Text, int size = 80)
        {
            EUCL.SetCursorPos(x, y, size);
            EUCL.SendStr(Text);
        }
        private string ClearWhiteSpacesStr(string text)
        {
            return text.Replace("  ", "");
        }
        public void CloseConnection()
        {
            EUCL.ClearScreen();
            EUCL.Disconnect(sessionName.ToString());
            Console.WriteLine("Соединение с Equation закрыто");
        }
    }
}
