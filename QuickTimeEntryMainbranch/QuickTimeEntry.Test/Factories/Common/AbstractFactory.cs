using Microsoft.Xrm.Sdk;

using A24.Xrm;

namespace A24.Xrm
{
    public abstract class AbstractFactory<T> where T : Entity
    {
        protected AbstractFactory()
        {
            CreateObject();
            FillRequiredAttributes();
            FillOptionalAttributes();
            FillCustomAttributes();
        }

        protected AbstractFactory(EntityReference erRef)
        {
            CreateObject();
        }

        public abstract T CurrentEntity { get; set; }

        public T GetNewObject()
        {
            CreateObject();
            FillRequiredAttributes();
            FillOptionalAttributes();
            FillCustomAttributes();

            return CurrentEntity;
        }
        
        protected abstract void CreateObject();

        protected abstract void FillRequiredAttributes();
        protected abstract void FillOptionalAttributes();
        protected abstract void FillCustomAttributes();
    }
}