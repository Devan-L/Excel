using CsvHelper;
using CsvHelper.Configuration;
using DelimitedFileSqlServerTableGenerator.Extensions;
using DelimitedFileSqlServerTableGenerator.Models;
using DelimitedFileSqlServerTableGenerator.Models.ParsingTypes;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

// TODO : Add conversion options to fields
// TODO : Move type resolution to resolver class

namespace DelimitedFileSqlServerTableGenerator.ViewModels
{
    [PropertyChanged.AddINotifyPropertyChangedInterface]
    internal class MainWindowViewModel
    {
        public string Delimeter { get; set; } = "\t";
        public bool HasHeaderRow { get; set; } = true;
        public string SelectedFilePath { get; set; }// = @"C:\Users\Samji\Downloads\MOCK_DATA (1).txt";
        public IEnumerable<object> Results { get; set; }
        [PropertyChanged.AlsoNotifyFor(nameof(SqlServerCreateStatement), nameof(SqlServerInsertStatement))]
        public string SchemaName { get; set; } = "dbo";
        [PropertyChanged.AlsoNotifyFor(nameof(SqlServerCreateStatement), nameof(SqlServerInsertStatement))]
        public string TableName { get; set; }

        [PropertyChanged.AlsoNotifyFor(nameof(SqlServerCreateStatement), nameof(SqlServerInsertStatement))]
        public IEnumerable<ParsingField> ParsingFields { get; set; }
        public string SqlServerCreateStatement
        {
            get
            {
                if (ParsingFields == null || string.IsNullOrWhiteSpace(SelectedFilePath))
                {
                    return null;
                }

                var columns = ParsingFields.Where(field => field.Include).Select(field => $"\t[{field.Name}] {field.SelectedType.SqlServerDataType} {(field.SelectedType.IsNullable ? "NULL" : "NOT NULL")}").ToList();

                var createTable = $"CREATE TABLE [{SchemaName}].[{TableName}] (\n{columns.Join(",\n")}\n)";
                return createTable;
            }
        }

        public string SqlServerInsertStatement
        {
            get
            {
                if (ParsingFields == null || string.IsNullOrWhiteSpace(SelectedFilePath))
                {
                    return null;
                }

                var inserts = Results.Cast<IDictionary<string, object>>().Select(row =>
                {
                    var columnNames = ParsingFields.Where(field => field.Include).Select(field => field.Name).Select(name => $"[{name}]").Join(",");
                    var values = ParsingFields.Where(field => field.Include).Select(field =>
                    {
                        var value = row[field.Name].ToString();
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            return "NULL";
                        }
                        if(field.SelectedType is DateTimeParsingType)
                        {
                            var dateValue = DateTime.Parse(value);
                            return $"'{dateValue:yyyy-MM-dd HH:mm:ss}'";
                        }
                        return $"'{value.Replace("'", "''")}'";
                    }).Join(",");
                    var insert = $"INSERT INTO [{SchemaName}].[{TableName}] ({columnNames}) VALUES ({values})";

                    return insert;
                });

                return inserts.Join(Environment.NewLine);
            }
        }

        [PropertyChanged.AlsoNotifyFor(nameof(SqlServerCreateStatement), nameof(SqlServerInsertStatement))]
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
                ParseFile();
            }
        }

        public void ParseFile()
        {
            try
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

                    this.Results = records;

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
                    var friendlyFileName = Regex.Replace(Path.GetFileNameWithoutExtension(SelectedFilePath), "[^a-zA-Z0-9]", "");
                    this.TableName = friendlyFileName;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                throw;
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