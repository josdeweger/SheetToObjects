using System;

namespace SheetToObjects.Lib.FluentConfiguration
{
    public abstract class SheetToObjectConfig
    {
        public MappingConfig MappingConfig { get; private set; }
        public Type ModelType { get; private set; }

        protected void CreateMap<TModel>(Func<MappingConfigBuilder<TModel>, MappingConfigBuilder<TModel>> mappingConfigFunc)
            where TModel : new()
        {
            ModelType = typeof(TModel);
            MappingConfig = mappingConfigFunc(new MappingConfigBuilder<TModel>()).Build();
        }
    }
}