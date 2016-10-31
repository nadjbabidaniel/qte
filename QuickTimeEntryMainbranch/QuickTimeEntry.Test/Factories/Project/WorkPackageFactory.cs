using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CCMA24Produktiv;

namespace A24.Xrm
{
    class WorkPackageFactory : AbstractFactory<a24_workpackage>
    {
        public override a24_workpackage CurrentEntity
        {
            get;
            set;
        }

        protected override void CreateObject()
        {
            CurrentEntity = new a24_workpackage();
        }

        protected override void FillRequiredAttributes()
        {
            CurrentEntity.a24_matchcode_str = "Test JNE " + (new Guid()).ToString();
        }

        protected override void FillOptionalAttributes()
        {
        }

        protected override void FillCustomAttributes()
        {
        }
    }
}
