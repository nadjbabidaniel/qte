using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using A24.Xrm;
using CCMA24Produktiv;

using Microsoft.Xrm.Sdk;

namespace QuickTimeEntry.Test
{
    [TestClass]
    public class RecordHelper
    {
        [TestMethod]
        public void CreateRequiredRecords()
        {
            var orgFactory = new OrganizationServiceFactory();
            var orgSrv = orgFactory.GetNewOrganizationService();
            var context = new XrmServiceContext(orgSrv);

            var accFac = new AccountFactory();
            context.AddObject(accFac.CurrentEntity);

            var conFac = new ContactFactory();
            context.AddRelatedObject(accFac.CurrentEntity, new Relationship("contact_customer_accounts"), conFac.CurrentEntity);

            context.SaveChanges();

            var incFac = new IncidentFactory();
            incFac.CurrentEntity.CustomerId = accFac.CurrentEntity.ToEntityReference();
            context.AddObject(incFac.CurrentEntity);
            context.Attach(accFac.CurrentEntity);

            var wpFac = new WorkPackageFactory();
            context.AddRelatedObject(accFac.CurrentEntity, new Relationship("a24_account_workpackage_customer_ref"), wpFac.CurrentEntity);
            wpFac.CurrentEntity.a24_customer_ref = accFac.CurrentEntity.ToEntityReference();

            context.SaveChanges();
        }
    }
}
