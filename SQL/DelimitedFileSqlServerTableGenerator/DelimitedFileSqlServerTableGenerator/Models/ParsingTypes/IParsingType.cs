namespace DelimitedFileSqlServerTableGenerator.Models.ParsingTypes
{
    internal interface IParsingType
    {
        public bool IsNullable { get; set; }
        public string SqlServerDataType { get; }
    }
}
