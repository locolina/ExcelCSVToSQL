using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCSVToSQL.Strings
{
    internal class Strings
    {
        public const string createTable = "CREATE TABLE [##{0}]({1})";
        public const string dropTable = "DROP TABLE IF EXISTS [##{0}]";
        public const string tableColumns = "[{0}] NVARCHAR(MAX)";
        public const string tableInsert = "INSERT INTO [##{0}]({1})";
        public const string insertColumns = "VALUES ({0})";
        public const string cellValue = "[{0}]";
        public const string cellItem = "N'{0}'";
    }
}
