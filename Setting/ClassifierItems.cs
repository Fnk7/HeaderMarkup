using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace HeaderMarkup.Setting
{
    class ClassifierItems : StringConverter
    {
        public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
        {
            return true;
        }

        public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
        {
            if (context != null && context.Instance is Settings settings)
            {
                var values = Directory.GetFiles(settings.PythonFiles, "*.pkl").Select(file =>
                {
                    return Path.GetFileName(file);
                });
                return new StandardValuesCollection(values.ToList());
            }
            return base.GetStandardValues(context);
        }

        public override bool GetStandardValuesExclusive(ITypeDescriptorContext context)
        {
            return false;
        }
    }
}
