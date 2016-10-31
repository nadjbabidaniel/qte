using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CCMA24Produktiv;

namespace A24.Xrm
{
    class IncidentFactory : AbstractFactory<Incident>
    {
        public override Incident CurrentEntity { get; set; }

        protected override void CreateObject()
        {
            this.CurrentEntity = new Incident();
        }

        protected override void FillRequiredAttributes()
        {
            CurrentEntity.Title = "Test JNE " + (new Guid()).ToString();
        }

        protected override void FillOptionalAttributes()
        {
        }

        protected override void FillCustomAttributes()
        {
        }
    }
}
