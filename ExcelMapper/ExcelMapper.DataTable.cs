using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Ganss.Excel
{
    public partial class ExcelMapper
    {

        /// <summary>
        /// Saves the DataTable to the specified Excel file.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(string file, DataTable dataTable, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
            Save(fs, dataTable, sheetName, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves the DataTable to the specified Excel file.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(string file, DataTable dataTable, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
            Save(fs, dataTable, sheetIndex, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves the DataTable to the specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the DataTable to.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(Stream stream, DataTable dataTable, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            this.IgnoreNestedTypes = true;
            Save(stream, ConverterToDynamic(dataTable), sheetName, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves the DataTable to the specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the DataTable to.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(Stream stream, DataTable dataTable, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            Save(stream, ConverterToDynamic(dataTable), sheetIndex, xlsx, valueConverter);
        }


        /// <summary>
        /// Saves the DataTable to the specified Excel file using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(string file, DataTable dataTable, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var ms = new MemoryStream();
            Save(ms, dataTable, sheetName, xlsx, valueConverter);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves the DataTable to the specified Excel file using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(string file, DataTable dataTable, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var ms = new MemoryStream();
            Save(ms, dataTable, sheetIndex, xlsx, valueConverter);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves the DataTable to the specified stream using async I/O.
        /// </summary>
        /// <param name="stream">The stream to save the DataTable to.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(Stream stream, DataTable dataTable, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var ms = new MemoryStream();
            Save(ms, dataTable, sheetName, xlsx, valueConverter);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Saves the DataTable to the specified stream using async I/O.
        /// </summary>
        /// <param name="stream">The stream to save the DataTable to.</param>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(Stream stream, DataTable dataTable, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var ms = new MemoryStream();
            Save(ms, dataTable, sheetIndex, xlsx, valueConverter);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Adds a mapping from a excel column name to a data column.
        /// </summary>s
        /// <param name="cols">The cols that contains the col to map to.</param>
        /// <param name="headerText">excel header text.</param>
        /// <param name="dataColumnName">Name of the property.</param>
        public ColumnInfo AddMapping(DataColumnCollection cols, string headerText, string dataColumnName)
        {
            if (TypeMapperFactory == DefaultTypeMapperFactory)
            {
                TypeMapperFactory = new DataColumnMapperFactory();
            }
            TypeMapper typeMapper = TypeMapperFactory.Create(ConverterToDynamic(cols));
            var dataColumn = cols[dataColumnName];

            if (!typeMapper.ColumnsByName.ContainsKey(headerText))
                typeMapper.ColumnsByName.Add(headerText, new List<ColumnInfo>());
            if (dataColumnName != headerText && typeMapper.ColumnsByName.ContainsKey(dataColumnName))
                typeMapper.ColumnsByName.Remove(dataColumnName);

            var columnInfo = typeMapper.ColumnsByName[headerText].FirstOrDefault(ci => ci.Name == dataColumnName);
            if (columnInfo is null)
            {
                columnInfo = new DynamicColumnInfo(dataColumnName, dataColumn.DataType.ConvertToNullableType());
                typeMapper.ColumnsByName[headerText].Add(columnInfo);
            }
            
            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a excel column index to a data column.
        /// </summary>
        /// <param name="cols">The cols that contains the col to map to.</param>
        /// <param name="headerIndex">Index of the excel column. min value is 1</param>
        /// <param name="dataColumnName">Name of the property.</param>
        public ColumnInfo AddMapping(DataColumnCollection cols, int headerIndex, string dataColumnName)
        {
            if (TypeMapperFactory == DefaultTypeMapperFactory)
            {
                TypeMapperFactory = new DataColumnMapperFactory();
            }
            TypeMapper typeMapper = TypeMapperFactory.Create(ConverterToDynamic(cols));
            var dataColumn = cols[dataColumnName];
            var idx = headerIndex - 1;

            if (!typeMapper.ColumnsByIndex.ContainsKey(idx))
                typeMapper.ColumnsByIndex.Add(idx, new List<ColumnInfo>());
            if (typeMapper.ColumnsByName.ContainsKey(dataColumnName))
                typeMapper.ColumnsByName.Remove(dataColumnName);

            var columnInfo = typeMapper.ColumnsByIndex[idx].FirstOrDefault(ci => ci.Property.Name == dataColumnName);
            if (columnInfo is null)
            {
                columnInfo = new DynamicColumnInfo(dataColumnName, dataColumn.DataType.ConvertToNullableType());
                typeMapper.ColumnsByIndex[idx].Add(columnInfo);
            }

            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a excel column name to a data column.
        /// </summary>s
        /// <param name="colType">The colType that contains the col to map to.</param>
        /// <param name="headerText">Name of the column.</param>
        /// <param name="dataColumnName">Name of the property.</param>
        public ColumnInfo AddMapping(DbType colType, string headerText, string dataColumnName)
        {
            if (TypeMapperFactory == DefaultTypeMapperFactory)
            {
                TypeMapperFactory = new DataColumnMapperFactory();
            }
            TypeMapper typeMapper = TypeMapperFactory.Create(typeof(ExpandoObject));

            if (dataColumnName != headerText && typeMapper.ColumnsByName.Keys.Any(n => n == dataColumnName))
                typeMapper.ColumnsByName.Remove(dataColumnName);

            if (!typeMapper.ColumnsByName.Keys.Any(n=>n == headerText))
                typeMapper.ColumnsByName.Add(headerText, new List<ColumnInfo>());

            var columnInfo = typeMapper.ColumnsByName[headerText].FirstOrDefault(ci => ci.Name == dataColumnName);
            if (columnInfo is null)
            {
                var dataType = NetType2DbTypeMapping.Where(n => n.Value == colType).FirstOrDefault().Key;
                columnInfo = new DynamicColumnInfo(dataColumnName, dataType.ConvertToNullableType());
                typeMapper.ColumnsByName[headerText].Add(columnInfo);
            }

            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a excel column index to a data column.
        /// </summary>
        /// <param name="colType">The colType that contains the col to map to.</param>
        /// <param name="headerIndex">Index of the excel column. min value is 1</param>
        /// <param name="dataColumnName">Name of the property.</param>
        public ColumnInfo AddMapping(DbType colType, int headerIndex, string dataColumnName)
        {
            if (TypeMapperFactory == DefaultTypeMapperFactory)
            {
                TypeMapperFactory = new DataColumnMapperFactory();
            }
            TypeMapper typeMapper = TypeMapperFactory.Create(typeof(ExpandoObject));
            var idx = headerIndex - 1;

            if (!typeMapper.ColumnsByIndex.ContainsKey(idx))
                typeMapper.ColumnsByIndex.Add(idx, new List<ColumnInfo>());
            if (typeMapper.ColumnsByName.ContainsKey(dataColumnName))
                typeMapper.ColumnsByName.Remove(dataColumnName);

            var columnInfo = typeMapper.ColumnsByIndex[idx].FirstOrDefault(ci => ci.Property.Name == dataColumnName);
            if (columnInfo is null)
            {
                var dataType = NetType2DbTypeMapping.Where(n => n.Value == colType).FirstOrDefault().Key;
                columnInfo = new DynamicColumnInfo(dataColumnName, dataType.ConvertToNullableType());
                typeMapper.ColumnsByIndex[idx].Add(columnInfo);
            }

            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a excel column name to a data column.
        /// </summary>s
        /// <param name="cols">The cols that contains the col to map to.</param>
        /// <param name="columnMapping">columnName and excel header text mapping table, Key is columnName, Value is header text</param>
        public List<ColumnInfo> AddMapping(DataColumnCollection cols, Dictionary<string, string> columnMapping)
        {
            var columnInfoList = new List<ColumnInfo>();
            foreach (var item in columnMapping)
            {
                columnInfoList.Add(this.AddMapping(cols.IndexOf(item.Key) != -1 ? NetType2DbTypeMapping.Where(n => n.Key == cols[item.Key].DataType).FirstOrDefault().Value : DbType.String, item.Value, item.Key));
            }
            return columnInfoList;
        }

        /// <summary>
        /// Ignores a property.
        /// </summary>
        /// <param name="cols">The cols that contains the data column to map to.</param>
        /// <param name="dataColumnName">Name of the data column.</param>
        public void Ignore(DataColumnCollection cols, string dataColumnName)
        {
            if (TypeMapperFactory == DefaultTypeMapperFactory)
            {
                TypeMapperFactory = new DataColumnMapperFactory();
            }
            TypeMapper typeMapper = TypeMapperFactory.Create(ConverterToDynamic(cols));
            typeMapper.ColumnsByName.Where(c => c.Value.Any(cc => cc.Name.Equals(dataColumnName, StringComparison.OrdinalIgnoreCase)))
                .ToList().ForEach(kvp => typeMapper.ColumnsByName.Remove(kvp.Key));

            var col = typeMapper.ColumnsByIndex.FirstOrDefault(c =>
                c.Value.Any(cc => cc.Name.Equals(dataColumnName, StringComparison.OrdinalIgnoreCase)));
            if (col.Key != -1)
            {
                var afterColumn = typeMapper.ColumnsByIndex.Where(n => n.Key > col.Key).ToDictionary(d => d.Key, d => d.Value);
                foreach (var index in typeMapper.ColumnsByIndex.Keys.Where(k => k >= col.Key))
                {
                    typeMapper.ColumnsByIndex.Remove(index);
                }
                //fix index number
                foreach (var index in afterColumn.Keys)
                {
                    typeMapper.ColumnsByIndex.Add(index - 1, afterColumn[index]);
                }
            }
        }


        /// <summary>
        /// Fetches DataTable from the specified sheet name.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when a sheet is not found</exception>
        public DataTable FetchToDataTable(string sheetName, Func<string, object, object> valueParser = null)
        {
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetName), sheetName, "Sheet not found");
            }
            return FetchToDataTable(sheet, valueParser);
        }

        /// <summary>
        /// Fetches DataTable from the specified sheet index.
        /// </summary>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public DataTable FetchToDataTable(int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return FetchToDataTable(sheet, valueParser);
        }

        DataTable FetchToDataTable(ISheet sheet, Func<string, object, object> valueParser = null)
        {
            var list = Fetch(sheet, typeof(ExpandoObject), valueParser).OfType<dynamic>();
            return ToDataTable(list);
        }

        public static List<dynamic> ConverterToDynamic(DataTable dataTable)
        {
            var list = new List<dynamic>();
            var columns = dataTable.Columns.Cast<DataColumn>();
            var map = columns.ToDictionary(c => c.ColumnName, c => c.Ordinal);
            foreach (DataRow row in dataTable.Rows)
            {
                var eo = new ExpandoObject();
                var expando = (IDictionary<string, object>)eo;
                expando[TypeMapper.IndexMapPropertyName] = map;
                foreach (DataColumn column in columns)
                {
                    expando[column.ColumnName] = row[column.ColumnName] == DBNull.Value ? null : row[column.ColumnName].ConvertToNullable();
                }
                list.Add(eo);
            }
            return list;
        }
        
        public static dynamic ConverterToDynamic(DataColumnCollection columns)
        {
            var map = columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => c.Ordinal);
            var eo = new ExpandoObject();
            var expando = (IDictionary<string, object>)eo;
            expando[TypeMapper.IndexMapPropertyName] = map;
            foreach (DataColumn column in columns)
            {
                expando[column.ColumnName] = GetDefaultValue(column.DataType).ConvertToNullable();
            }
            return eo;
        }

        public static DataTable ToDataTable(IEnumerable<dynamic> items)
        {
            DataTable dtDataTable = new DataTable();
            if (items.Count() == 0) return dtDataTable;

            ((IDictionary<string, object>)items.First()).ToList().ForEach(col => { dtDataTable.Columns.Add(col.Key, col.Value.GetType()); });

            foreach (IDictionary<string, object> item in items)
            {
                DataRow dr = dtDataTable.NewRow();
                item.ToList().ForEach(Col => { dr[Col.Key] = Col.Value; });
                dtDataTable.Rows.Add(dr);
            }
            return dtDataTable;
        }        
        /// <summary>
        /// Initialize default values based on type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static object GetDefaultValue(Type type)
        {
            if (type == typeof(string))
            {
                return string.Empty;
            }
            else if (type == typeof(int))
            {
                return 0;
            }
            else if (type == typeof(long))
            {
                return 0L;
            }
            else if (type == typeof(double))
            {
                return 0.0;
            }
            else if (type == typeof(decimal))
            {
                return 0M;
            }
            else if (type == typeof(bool))
            {
                return false;
            }
            else if (type == typeof(DateTime))
            {
                return DateTime.MinValue;
            }
            else if (type == typeof(Guid))
            {
                return Guid.Empty;
            }
            else if (type == typeof(byte[]))
            {
                return new byte[0];
            }
            else if (type == typeof(char))
            {
                return '\0';
            }
            else if (type == typeof(short))
            {
                return (short)0;
            }
            else if (type == typeof(ushort))
            {
                return (ushort)0;
            }
            else if (type == typeof(uint))
            {
                return (uint)0;
            }
            else if (type == typeof(ulong))
            {
                return (ulong)0;
            }
            else if (type == typeof(sbyte))
            {
                return (sbyte)0;
            }
            else if (type == typeof(float))
            {
                return (float)0.0;
            }
            else if (type == typeof(TimeSpan))
            {
                return TimeSpan.MinValue;
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// .NET data type and DbType mapping
        /// </summary>
        public static List<KeyValuePair<Type, DbType>> NetType2DbTypeMapping
        {
            get
            {
                List<KeyValuePair<Type, DbType>> map = new List<KeyValuePair<Type, DbType>>();
                map.Add(new KeyValuePair<Type, DbType>(typeof(object), DbType.Object));
                map.Add(new KeyValuePair<Type, DbType>(typeof(byte), DbType.Byte));
                map.Add(new KeyValuePair<Type, DbType>(typeof(sbyte), DbType.SByte));
                map.Add(new KeyValuePair<Type, DbType>(typeof(short), DbType.Int16));
                map.Add(new KeyValuePair<Type, DbType>(typeof(ushort), DbType.UInt16));
                map.Add(new KeyValuePair<Type, DbType>(typeof(int), DbType.Int32));
                map.Add(new KeyValuePair<Type, DbType>(typeof(uint), DbType.UInt32));
                map.Add(new KeyValuePair<Type, DbType>(typeof(long), DbType.Int64));
                map.Add(new KeyValuePair<Type, DbType>(typeof(ulong), DbType.UInt64));
                map.Add(new KeyValuePair<Type, DbType>(typeof(float), DbType.Single));
                map.Add(new KeyValuePair<Type, DbType>(typeof(double), DbType.Double));
                map.Add(new KeyValuePair<Type, DbType>(typeof(decimal), DbType.Decimal));
                map.Add(new KeyValuePair<Type, DbType>(typeof(bool), DbType.Boolean));
                map.Add(new KeyValuePair<Type, DbType>(typeof(string), DbType.String));
                map.Add(new KeyValuePair<Type, DbType>(typeof(char), DbType.StringFixedLength));
                map.Add(new KeyValuePair<Type, DbType>(typeof(Guid), DbType.Guid));
                map.Add(new KeyValuePair<Type, DbType>(typeof(DateTime), DbType.DateTime));
                map.Add(new KeyValuePair<Type, DbType>(typeof(DateTime), DbType.DateTime2));
                map.Add(new KeyValuePair<Type, DbType>(typeof(DateTimeOffset), DbType.DateTimeOffset));
                map.Add(new KeyValuePair<Type, DbType>(typeof(byte[]), DbType.Binary));
                map.Add(new KeyValuePair<Type, DbType>(typeof(string), DbType.Xml));

                return map;
            }
        }
    }
}
