using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrationPDFGeneration
{
    class CommonDefinitions
    {
        public class PDFStream
        {
            public Byte[] stream;
            public String fileName;
            public String identifier;
           

            public PDFStream(Byte[] stream, String fileName, String identifier = "")
            {
                this.stream = stream;
                this.fileName = fileName;
                this.identifier = identifier;
            }
        }

        public class multiDocument
        {
            public String url;
            public String identifier;

            public multiDocument(String url, String identifier = "")
            {
                this.url = url;
                this.identifier = identifier;
            }
        }


        public class ColumnDefinition
        {
            public String columnName;
            public String fieldName;
            public String defaultValue = " ";
            public Boolean isTotalColumn = false;
            public int columnWidth;

            public ColumnDefinition(String columnName, String fieldName)
            {
                this.columnName = columnName;
                this.fieldName = fieldName;
            }
            public ColumnDefinition(String columnName, String fieldName, String columnWidth, String defaultValue ="")
            {
                this.columnName = columnName;
                this.fieldName = fieldName;
                this.defaultValue = defaultValue;
                if (this.defaultValue == "0")
                {
                    this.isTotalColumn = true;
                }

                int.TryParse(columnWidth, out this.columnWidth);
            }
        }
        public class TotalColumns
        {
            public String tableId;
            public String fieldName;
            public Double total;
            public int fieldIndex;

            public TotalColumns(String tableId, int fieldIndex, String fieldName)
            {
                this.tableId = tableId;
                this.fieldName = fieldName;
                this.total = 0;
                this.fieldIndex = fieldIndex;
            }
        }
    }
}
