using System;

namespace DelimitedFileSqlServerTableGenerator.Models.ParsingTypes
{
    internal abstract class ParsingType<T> : IParsingType
    {
        public bool IsNullable { get; set; }

        private string sqlServerDataType;
        public string SqlServerDataType
        {
            get { return sqlServerDataType ?? DefaultSqlServerDataType; }
            set { sqlServerDataType = value; }
        }

        public abstract string DefaultSqlServerDataType { get; }
    }

    // TODO: Move to own files
    internal class StringParsingType : ParsingType<string>
    {
        public int Length { get; set; }

        public override string DefaultSqlServerDataType => $"VARCHAR({Length})";
    }

    internal class IntegerParsingType : ParsingType<int>
    {
        public override string DefaultSqlServerDataType => "INT";
    }

    internal class DecimalParsingType : ParsingType<decimal>
    {
        public int Precision { get; set; }

        public override string DefaultSqlServerDataType => $"DECIMAL(0, {Precision})";
    }

    internal class DateTimeParsingType : ParsingType<DateTime>
    {
        public override string DefaultSqlServerDataType => "DATETIME";
    }

    internal class BooleanParsingType : ParsingType<DateTime>
    {
        public override string DefaultSqlServerDataType => "BIT";
    }
}
