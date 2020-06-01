using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace DelimitedFileSqlServerTableGenerator.ViewModels
{
    [PropertyChanged.AddINotifyPropertyChangedInterface]
    public class MainWindowViewModel
    {
        public string Delimeter { get; set; } = "\t";
        public bool HasHeaderRow { get; set; } = true;
        public string SelectedFilePath { get; set; } = @"C:\Users\Samji\Downloads\MOCK_DATA (1).txt";
        public IEnumerable<object> Results { get; set; }

        [PropertyChanged.AlsoNotifyFor(nameof(SqlServerCreateStatement))]
        public IEnumerable<ParsingField> ParsingFields { get; set; }
        public string SqlServerCreateStatement
        {
            get
            {
                if (ParsingFields == null || string.IsNullOrWhiteSpace(SelectedFilePath))
                {
                    return null;
                }

                var friendlyFileName = Regex.Replace(Path.GetFileNameWithoutExtension(SelectedFilePath), "[^a-zA-Z]", "");

                var columns = ParsingFields.Select(field => $"\t[{field.Name}] {field.SelectedType.SqlServerDataType} {(field.SelectedType.IsNullable ? "NULL" : "NOT NULL")}").ToList();

                var createTable = $"CREATE TABLE [{friendlyFileName}] (\n{columns.Join(",\n")}\n)";
                return createTable;
            }
        }

        [PropertyChanged.AlsoNotifyFor(nameof(SqlServerCreateStatement))]
        private DateTime refreshSqlStatement { get; set; }

        public void RefreshSqlServerCreateStatement()
        {
            refreshSqlStatement = DateTime.Now;
        }

        public void SelectFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                CheckFileExists = true,
                Multiselect = false
            };

            if (openFileDialog.ShowDialog().GetValueOrDefault(false))
            {
                this.SelectedFilePath = openFileDialog.FileName;
            }
        }

        public void ParseFile()
        {
            var csvConfiguration = new CsvConfiguration(CultureInfo.CurrentCulture)
            {
                Delimiter = Delimeter,
                HasHeaderRecord = HasHeaderRow
            };

            using var reader = new StreamReader(SelectedFilePath);
            using (var csv = new CsvReader(reader, csvConfiguration))
            {
                var records = csv.GetRecords<dynamic>().ToList();

                //this.Results = records;

                var recordsAsDictionary = records.Cast<IDictionary<string, object>>().ToList();

                var properties = recordsAsDictionary.FirstOrDefault();

                var fields = properties.Keys.Select(property =>
                {
                    var valuesForProperty = recordsAsDictionary.Select(record => record[property]).Cast<string>().ToList();
                    var applicableTypes = EvaluateType(valuesForProperty).ToList();
                    return new ParsingField
                    {
                        Name = property,
                        ApplicableTypes = applicableTypes,
                        SelectedType = applicableTypes.FirstOrDefault()
                    };
                }).ToList();

                this.ParsingFields = fields;
            }
        }

        private IEnumerable<IParsingType> EvaluateType(IEnumerable<string> values)
        {
            var applicableTypes = new List<IParsingType>();

            var populatedProperties = values.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();

            var isNullable = values.Any(value => string.IsNullOrWhiteSpace(value));

            // Numeric
            if (IsInteger(populatedProperties))
            {
                applicableTypes.Add(new IntegerParsingType { IsNullable = isNullable });
            }

            if (IsDecimal(populatedProperties))
            {
                applicableTypes.Add(new DecimalParsingType { IsNullable = isNullable, Precision = 6 });
            }

            // Boolean
            // Y/N, Yes/No, True/False, 1/0
            if (IsBoolean(populatedProperties))
            {
                applicableTypes.Add(new BooleanParsingType { IsNullable = isNullable });
            }

            // Date 
            // TODO : Date or time ?
            if (IsDate(populatedProperties))
            {
                applicableTypes.Add(new DateTimeParsingType { IsNullable = isNullable });
            }

            // String
            var maxLength = values.Max(value => value.Length);
            applicableTypes.Add(new StringParsingType { IsNullable = isNullable, Length = maxLength });

            // BLOB?

            return applicableTypes;
        }

        private bool IsInteger(IEnumerable<object> values)
        {
            return values.All(value => int.TryParse(value.ToString(), out var result));
        }

        private bool IsBoolean(IEnumerable<object> values)
        {
            return values.All(value => bool.TryParse(value.ToString(), out var result));
        }

        private bool IsDecimal(IEnumerable<object> values)
        {
            return values.All(value => decimal.TryParse(value.ToString(), out var result));
        }

        private bool IsDate(IEnumerable<object> values)
        {
            return values.All(value => DateTime.TryParse(value.ToString(), out var result));
        }
    }
}

internal static class TypeExtensions
{
    internal static bool Is<T>(this Type input)
    {
        return input == typeof(T);
    }
}

internal static class IEnumerableExtensions
{
    internal static string Join<T>(this IEnumerable<T> collection, string separator)
    {
        return string.Join(separator, collection);
    }
}

public class ParsingField
{
    public string Name { get; set; }
    public IEnumerable<IParsingType> ApplicableTypes { get; set; }
    public IParsingType SelectedType { get; set; }
}


public interface IParsingType
{
    public bool IsNullable { get; set; }
    public string SqlServerDataType { get; }
}

public abstract class ParsingType<T> : IParsingType
{
    public bool IsNullable { get; set; }
    public abstract string SqlServerDataType { get; }
}

public class StringParsingType : ParsingType<string>
{
    public int Length { get; set; }

    public override string SqlServerDataType => $"VARCHAR({Length})";
}

public class IntegerParsingType : ParsingType<int>
{
    public override string SqlServerDataType => "INT";
}

public class DecimalParsingType : ParsingType<decimal>
{
    public int Precision { get; set; }

    public override string SqlServerDataType => $"DECIMAL(0, {Precision})";
}

public class DateTimeParsingType : ParsingType<DateTime>
{
    public override string SqlServerDataType => "DATETIME";
}

public class BooleanParsingType : ParsingType<DateTime>
{
    public override string SqlServerDataType => "BIT";
}
