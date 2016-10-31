//using A24.Xrm.Helper;
using A24.Xrm;
using System;

using CCMA24Produktiv;

namespace A24.Xrm
{
    class ContactFactory : AbstractFactory<Contact>
    {
        public override Contact CurrentEntity { get; set; }

        protected override void CreateObject()
        {
            CurrentEntity = new Contact();
        }

        protected override void FillRequiredAttributes()
        {
            //CurrentEntity.FirstName = new XMLDataHelper().getXmlData(Constants.XmlFileNames.FirstName, "FirstName");
            CurrentEntity.FirstName = String.Format("JNE {0}", Guid.NewGuid().ToString());
            //CurrentEntity.LastName = new XMLDataHelper().getXmlData(Constants.XmlFileNames.LastName, "LastName");
            CurrentEntity.LastName = String.Format("JNE {0}", Guid.NewGuid().ToString());
        }

        protected override void FillOptionalAttributes()
        {
            //CurrentEntity.Address1_Name = new XMLDataHelper().getXmlData(Constants.XmlFileNames.StreetName, "StreetName");
        }

        protected override void FillCustomAttributes()
        {
            CurrentEntity.Attributes = CurrentEntity.Attributes;
        }
    }
}
