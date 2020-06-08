using System;
using System.Collections.Generic;
using IBM.Broker.Plugin;
using System.Linq;
using System.Data;

namespace IntegrationPDFGeneration
{
    class RTFBuilder
    {
        public String tableHeader = @"{\rtlch\fcs1 \af31507 \ltrch\fcs0 \insrsid2712186 \par \ltrrow}\trowd \irow0\irowband0\ltrrow\ts15\trgaph57\trleft0\trbrdrt\brdrs\brdrw10\brdrcf19 \trbrdrl\brdrs\brdrw10\brdrcf19 \trbrdrb\brdrs\brdrw10\brdrcf19 \trbrdrr\brdrs\brdrw10\brdrcf19 \trbrdrh\brdrs\brdrw10\brdrcf19 \trbrdrv\brdrs\brdrw10\brdrcf19 \trftsWidth3\trwWidth10485\trftsWidthB3\trftsWidthA3\trpaddl57\trpaddt142\trpaddb142\trpaddr57\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid9778608\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind0\tblindtype3 \clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth704\clcbpatraw20 \cellx704\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth2268\clcbpatraw20 \cellx2972\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth3402\clcbpatraw20 \cellx6374\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth567\clcbpatraw20 \cellx6941\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth1276\clcbpatraw20 \cellx8217\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth709\clcbpatraw20 \cellx8926\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth1559\clcbpatraw20 \cellx10485\pard\plain \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid1393303\yts15 \rtlch\fcs1 \af31507\afs24\alang1025 \ltrch\fcs0 \f1\fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033";
        private String tableRow = @"\pard\plain \ltrpar\ql \li0\ri0\sa200\sl276\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin[rowid] \rtlch\fcs1 \af31507\afs24\alang1025 \ltrch\fcs0 \f1\fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\rtlch\fcs1 \af31507 \ltrch\fcs0 \insrsid5657135\charrsid14622050 \trowd \irow[rowid]\irowband[rowid]\ltrrow\ts15\trgaph57\trleft0\trbrdrt\brdrs\brdrw10\brdrcf19 \trbrdrl\brdrs\brdrw10\brdrcf19 \trbrdrb\brdrs\brdrw10\brdrcf19 \trbrdrr\brdrs\brdrw10\brdrcf19 \trbrdrh\brdrs\brdrw10\brdrcf19 \trbrdrv\brdrs\brdrw10\brdrcf19\trftsWidth3\trwWidth10485\trftsWidthB3\trftsWidthA3\trpaddl57\trpaddt142\trpaddb142\trpaddr57\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid9778608\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind0\tblindtype3 \clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth704\clcbpatraw20 \cellx704\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone\clcbpat20\cltxlrtb\clftsWidth3\clwWidth2268\clcbpatraw20 \cellx2972\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth3402\clcbpatraw20 \cellx6374\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth567\clcbpatraw20 \cellx6941\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone\clcbpat20\cltxlrtb\clftsWidth3\clwWidth1276\clcbpatraw20 \cellx8217\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth709\clcbpatraw20 \cellx8926\clvertalt\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \clcbpat20\cltxlrtb\clftsWidth3\clwWidth1559\clcbpatraw20 \cellx10485\row \ltrrow}\trowd ";
        private String tableFooter = @"{\irow[rowid]\irowband1\lastrow \ltrrow\ts15\trgaph57\trleft0\trbrdrt\brdrs\brdrw10\brdrcf19 \trbrdrl\brdrs\brdrw10\brdrcf19\trbrdrb\brdrs\brdrw10\brdrcf19 \trbrdrr\brdrs\brdrw10\brdrcf19 \trbrdrh\brdrs\brdrw10\brdrcf19 \trbrdrv\brdrs\brdrw10\brdrcf19\trftsWidth3\trwWidth10485\trftsWidthB3\trftsWidthA3\trpaddl57\trpaddt142\trpaddb142\trpaddr57\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5657135\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind0\tblindtype3 \clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19\clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth704\clshdrawnil \cellx704\clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19 \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone\cltxlrtb\clftsWidth3\clwWidth2268\clshdrawnil \cellx2972\clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19 \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3402\clshdrawnil \cellx6374\clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19 \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth567\clshdrawnil \cellx6941\clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19 \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth1276\clshdrawnil \cellx8217\clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19 \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth709\clshdrawnil \cellx8926\clvertalt\clbrdrt\brdrs\brdrw10\brdrcf19 \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw10\brdrcf19 \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth1559\clshdrawnil \cellx10485\row }\pard \ltrpar\ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0\pararsid3898411";
        private String tableDataCell = @"{\rtlch\fcs1 \af31507 \ltrch\fcs0 \insrsid1210128 [[DATA]]}{\rtlch\fcs1 \af31507 \ltrch\fcs0 \insrsid1271067\charrsid16199972 \cell }";
        private String tableHeaderCell = @"{\rtlch\fcs1 \af31507 \ltrch\fcs0 \b\insrsid1271067\charrsid14622050 [[DATA]]\cell }";

        private String tableDataRowEnd = @"\pard\plain \ltrpar\ql \li0\ri0\sa200\sl276\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 \rtlch\fcs1 \af31507\afs24\alang1025 \ltrch\fcs0\f1\fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\rtlch\fcs1 \af31507 \ltrch\fcs0 \insrsid5657135\charrsid16199972 \trowd";
        private String tableDataRowEndLast = @" \trowd ";


        public String addHeaderCell(String columnHeaderString)
        {
            String headerCell = tableHeaderCell;
            headerCell = headerCell.Replace("[[DATA]]", columnHeaderString);

            return headerCell;
        }

        public String addTableFoolder(int rowCount)
        {
            return tableFooter.Replace("[rowid]", (rowCount).ToString());
        }

        public string addTableRow(int rowid, List<CommonDefinitions.ColumnDefinition> colDefs, NBElement rowData, Boolean isLastRow, List<CommonDefinitions.TotalColumns> totalCols)
        {
            string rowString = tableRow.Replace("[rowid]", (rowid + 1).ToString());
            rowString += addTableDataRow(rowid, colDefs, rowData, isLastRow, totalCols);
            return rowString;
        }

        public string addTableDataRow(int rowid, List<CommonDefinitions.ColumnDefinition> colDefs, NBElement rowData, Boolean isLastRow, List<CommonDefinitions.TotalColumns> totalCols)
        {
            String rowString = "";
            foreach (CommonDefinitions.ColumnDefinition colDef in colDefs)
            {
                if (colDef.fieldName == "lineNo")
                    rowString += tableDataCell.Replace("[[DATA]]", (rowid).ToString());
                else
                {
                    try
                    {
                        if (colDef.fieldName.StartsWith("CAL"))
                        {
                            String colFieldName = colDef.fieldName.Replace("CAL", "");
                            foreach (CommonDefinitions.ColumnDefinition cd in colDefs.Where(cd => cd.fieldName != "lineNo"))
                            {
                                if (rowData[cd.fieldName] != null)
                                    colFieldName = colFieldName.Replace("[" + cd.fieldName + "]", rowData[cd.fieldName].ValueAsString);
                            }

                            DataTable dt = new DataTable();
                            var val = dt.Compute(colFieldName, "");

                            rowString += tableDataCell.Replace("[[DATA]]", val.ToString());

                            addTotal(totalCols, colDef.fieldName, val);
                        }
                        else
                        {
                            if (rowData[colDef.fieldName] != null)
                            {
                                rowString += tableDataCell.Replace("[[DATA]]", rowData[colDef.fieldName].ValueAsString);
                                addTotal(totalCols, colDef.fieldName, rowData[colDef.fieldName].ValueAsString);
                            }
                            else
                                rowString += tableDataCell.Replace("[[DATA]]", colDef.defaultValue);
                        }
                    }
                    catch (Exception e)
                    {
                        rowString += tableDataCell.Replace("[[DATA]]", colDef.defaultValue);
                    }
                }
            }

            //rowString += "}";
            if (isLastRow == false)
                rowString += tableDataRowEnd;
            else
                rowString += tableRow.Replace("[rowid]", (rowid).ToString());

            return rowString;
        }

        private void addTotal(List<CommonDefinitions.TotalColumns> totalCols, String fieldName, Object val)
        {
            if (totalCols.Where(tc => tc.fieldName == fieldName).Count() > 0)
            {
                totalCols.Where(tc => tc.fieldName == fieldName).First().total = totalCols.Where(tc => tc.fieldName == fieldName).First().total + Convert.ToDouble(val);
            }
        }
    }
}
