using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OFKO_Robot.Model
{
    class FL_data
    {
        public string Pin; //Пин клиента
        public string Type; //Тип клиента
        public string Identification; //Идентификация
        public string ApproveType; // способ подтвреждения
        public string FIO; // ФИО
        public DateTime BirthDate;
        public string mnemonicFull;
        public string mnemonicShort;
        public string regCountry;
        public string Inn;
        public string SNILS;
        public string IsClient;
        public DateTime BecomeContragentDate;
        public DateTime BecomeClientDate;
        public string DocTypeInt;
        public string DocTypeStr;
        public string DocSeries;
        public string DocNumber;
        public string DocGiver;
        public DateTime DocOpenDate;
        public DateTime DocClosedDate;
        public string DocKPPCode;
        public string DocResponsibleUnit;
        public string DocResponsibleUnitFull;
        public string Citizenship;
        public string WorkPhone;
        public string MobilePhone;
        public string Email;
        public string Reserve1;
        public string Reserve2;
        public string AnyBeneficiary;
        public string PDL;
        public string BeneficiaryOwner;
        public string Reputation;
        public string FinancialPosition;
        public DateTime ActualizationDate;
        public string ActualizationEmployee;
        public string UpdateType;
        public DateTime LastUpdatedDate;
        public string UpdateEmployee;
        public string IsCrossedTheBorder;
        public DateTime CrossBorderDate;
        public string ResidenceDocTypeInt;
        public string ResidenceDocTypeStr;
        public string ResidenceSerieAndNumber;
        public DateTime ResidenceDateStart;
        public DateTime ResidenceDateEnd;
        public string MigrationCardTypeInt;
        public string MigrationCardNumber;
        public DateTime MigrationCardDateStart;
        public DateTime MigrationCardDateEnd;
        public int RiskLevelInt;
        public string RiskLevelStr;
        public List<Account> Accounts;
        public string RegistrationAddress;
        public string RegistrationAddressFull;
        public DateTime RegistrationUpdateDate;
        public string RegistrationPhone;
        public string FactAddress;
        public DateTime FactAddressUpdateDate;
        public string FactAddressPhone;
        public string RecordStatus;
        public string BirthPlace;
        public string BirthPlaceCountryCode;
        public string BirthPlaceCountry;
        public string BirthPlaceRegionCode;
        public string BirthPlaceRegionType;
        public string BirthPlaceRegion;
        public string BirthPlaceCityType;
        public string BirthPlaceCity;
        public string BirthPlaceLocalityType;
        public string BirthPlaceLocality;
        public string BirthPlaceFull;
        public string PrimePhone;
    }
}
