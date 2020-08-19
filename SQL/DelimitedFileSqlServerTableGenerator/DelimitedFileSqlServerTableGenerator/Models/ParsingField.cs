using DelimitedFileSqlServerTableGenerator.Models.ParsingTypes;
using System.Collections.Generic;

namespace DelimitedFileSqlServerTableGenerator.Models
{
    internal class ParsingField
    {
        public string Name { get; set; }
        public bool Include { get; set; } = true;
        public IEnumerable<IParsingType> ApplicableTypes { get; set; }
        public IParsingType SelectedType { get; set; }
    }
}
