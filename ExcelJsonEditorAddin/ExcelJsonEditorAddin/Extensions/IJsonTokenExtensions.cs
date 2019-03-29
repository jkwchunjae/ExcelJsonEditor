using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelJsonEditorAddin.JsonTokenModel;

namespace ExcelJsonEditorAddin
{
    public static class IJsonTokenExtensions
    {
        public static bool CanSpreadType(this JsonTokenType tokenType)
        {
            var spreadTypes = new[]
            {
                JsonTokenType.Object,
                JsonTokenType.Array,
            };

            return spreadTypes.Contains(tokenType);
        }

        public static bool CanSpreadType(this IJsonToken jsonToken)
            => jsonToken.Type().CanSpreadType();
    }
}
