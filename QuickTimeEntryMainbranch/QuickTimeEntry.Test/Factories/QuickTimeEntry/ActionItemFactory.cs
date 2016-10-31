using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using A24.Xrm;
using CCMA24Produktiv;

namespace A24.Xrm
{
    class ActionItemFactory : AbstractFactory<a24_action_item>
    {
        public override a24_action_item CurrentEntity { get; set; }

        protected override void CreateObject()
        {
            CurrentEntity = new a24_action_item();
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
