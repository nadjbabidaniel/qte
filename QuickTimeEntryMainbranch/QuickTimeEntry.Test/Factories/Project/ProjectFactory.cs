using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using A24.Xrm;
using CCMA24Produktiv;

namespace A24.Xrm
{
    class ProjectFactory : AbstractFactory<a24_project>
    {
        public override a24_project CurrentEntity {get;set;}

        protected override void CreateObject()
        {
            CurrentEntity = new a24_project();
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
