using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public enum JsonTokenType
    {
        Array,
        ObjectArray,
        Object,
        Property,
        Title,
        String,
        Number,
        Boolean,
        Other,
    }
}
