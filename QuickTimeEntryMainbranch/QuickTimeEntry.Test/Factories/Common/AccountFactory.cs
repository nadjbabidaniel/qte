using System;
//using A24.Xrm.Helper;
using A24.Xrm;

using CCMA24Produktiv;

namespace A24.Xrm
{
    public class AccountFactory : A24.Xrm.AbstractFactory<Account>
    {
        public override Account CurrentEntity { get; set; }

        protected override void CreateObject()
        {
            CurrentEntity = new Account();
        }

        protected override void FillRequiredAttributes()
        {
            //var accountNamePart1 = new XMLDataHelper().getXmlData(Constants.XmlFileNames.AccountName1, "AccountName1");
            //var accountNamePart2 = new XMLDataHelper().getXmlData(Constants.XmlFileNames.AccountName2, "AccountName2");
            //var accountNamePart3 = new XMLDataHelper().getXmlData(Constants.XmlFileNames.AccountName3, "AccountName3");
            //CurrentEntity.Name = String.Format("{0} {1} {2}", accountNamePart1, accountNamePart2, accountNamePart3);

            CurrentEntity.Name = String.Format("JNE {0}", Guid.NewGuid().ToString());
        }

        protected override void FillOptionalAttributes()
        {   
            //var streetName = new XMLDataHelper().getXmlData(Constants.XmlFileNames.StreetName, "StreetName");
            var streetName = CurrentEntity.Name = Guid.NewGuid().ToString();
            Random houseNumber = new Random();
            CurrentEntity.Address1_Name = string.Format("{0}, {1}", streetName, houseNumber.Next(500));
        }

        protected override void FillCustomAttributes()
        {
 
        }
    }
}
