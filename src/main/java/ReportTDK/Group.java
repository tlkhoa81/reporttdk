package ReportTDK;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Vector;
import jxl.*;
import jxl.biff.formula.FormulaException;
import jxl.format.CellFormat;
import jxl.format.Format;
import jxl.write.*;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */
class Group
{

    private WritableSheet sheet;
    private ResultSet rs;
    private String strGroupFieldName;
    private String strPGroupValue;
    private String strGroupValue;
    public int intPRow;
    public int intCRow;
    private int intDetailX;
    private int intTemplateRow;
    private int intColCount;
    private Cell vtTemplateCell[];
    public int vtCellType[];
    private Vector vtTemplateMerge;
    private Vector vtTemplateFormula;
    private Vector vtTemplateFormula1;
    private Vector vtFieldList[];
    private String vtFieldValueList[];
    private int intGroupType;
    private Total footer;

    public static String replace(String strSource, String strPattern, String strReplace)
    {
        if(strSource != null)
        {
            int intLen = strPattern.length();
            StringBuffer sb = new StringBuffer();
            int intFound = -1;
            int intStart;
            for(intStart = 0; (intFound = strSource.indexOf(strPattern, intStart)) != -1; intStart = intFound + intLen)
            {
                sb.append(strSource.substring(intStart, intFound));
                sb.append(strReplace);
            }

            sb.append(strSource.substring(intStart));
            return sb.toString();
        } else
        {
            return "";
        }
    }

    static void addMergeToSheet(WritableSheet sheet, Vector vtTemplateMerge, int intRow)
        throws WriteException
    {
        for(int intx = 0; intx < vtTemplateMerge.size(); intx++)
        {
            Merge merge1 = (Merge)vtTemplateMerge.get(intx);
            sheet.mergeCells(merge1.getC1(), intRow, merge1.getC2(), intRow);
        }

    }

    static void addMergeToSheet(WritableSheet sheet, Vector vtTemplateMerge)
        throws WriteException
    {
        for(int intx = 0; intx < vtTemplateMerge.size(); intx++)
        {
            Merge merge1 = (Merge)vtTemplateMerge.get(intx);
            sheet.mergeCells(merge1.getC1(), merge1.getR1(), merge1.getC2(), merge1.getR2());
        }

    }

    static String getCellValueToFill(ResultSet rs, Vector vtFieldList, String strCellContent)
        throws SQLException
    {
        int intFieldCount = vtFieldList.size();
        if(!strCellContent.equals("") && intFieldCount > 0)
        {
            for(int inti = 0; inti < intFieldCount; inti++)
            {
                String strFieldName = vtFieldList.get(inti).toString();
                String strFieldValue = "" + rs.getString(strFieldName);
                if(strFieldValue.equalsIgnoreCase("null"))
                {
                    strFieldValue = "";
                }
                strCellContent = replace(strCellContent, "${" + strFieldName + "}", strFieldValue);
            }

        }
        return strCellContent;
    }

    static boolean FieldExists(ResultSet rs, String strFieldName)
        throws SQLException
    {
        if(rs == null || strFieldName.equals(""))
        {
            return false;
        }

        return rs.findColumn(strFieldName) >= 1;
    }

    public Group(WritableSheet sheet, ResultSet rs, String strGroupFieldName, int intTemplateRow, int intColCount, int intDetailX)
        throws FormulaException, Exception
    {
        this.sheet = null;
        this.rs = null;
        this.strGroupFieldName = "";
        strPGroupValue = "(--)";
        strGroupValue = "";
        intPRow = 0;
        intCRow = 0;
        this.intDetailX = 0;
        this.intTemplateRow = 0;
        this.intColCount = 0;
        vtTemplateCell = null;
        vtCellType = null;
        vtTemplateMerge = null;
        vtTemplateFormula = null;
        vtTemplateFormula1 = null;
        vtFieldList = null;
        vtFieldValueList = null;
        intGroupType = 0;
        footer = null;
        intTemplateRow--;
        this.sheet = sheet;
        this.rs = rs;
        this.strGroupFieldName = strGroupFieldName;
        this.intTemplateRow = intTemplateRow;
        intCRow = intTemplateRow;
        this.intDetailX = intDetailX;
        this.intColCount = intColCount;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount];
        vtCellType = new int[intColCount];
        vtTemplateMerge = new Vector();
        vtTemplateFormula = new Vector();
        vtTemplateFormula1 = new Vector();
        vtFieldValueList = new String[intColCount];
        intGroupType = 1;
        parseGroupTemplate();
    }

    public Group(WritableSheet sheet, ResultSet rs, int intTemplateRow, int intColCount, int intDetailX)
        throws FormulaException, Exception
    {
        this.sheet = null;
        this.rs = null;
        strGroupFieldName = "";
        strPGroupValue = "(--)";
        strGroupValue = "";
        intPRow = 0;
        intCRow = 0;
        this.intDetailX = 0;
        this.intTemplateRow = 0;
        this.intColCount = 0;
        vtTemplateCell = null;
        vtCellType = null;
        vtTemplateMerge = null;
        vtTemplateFormula = null;
        vtTemplateFormula1 = null;
        vtFieldList = null;
        vtFieldValueList = null;
        intGroupType = 0;
        footer = null;
        intTemplateRow--;
        this.sheet = sheet;
        this.rs = rs;
        this.intTemplateRow = intTemplateRow;
        intCRow = intTemplateRow;
        this.intDetailX = intDetailX;
        this.intColCount = intColCount;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount];
        vtCellType = new int[intColCount];
        vtTemplateMerge = new Vector();
        vtTemplateFormula = new Vector();
        vtTemplateFormula1 = new Vector();
        vtFieldValueList = new String[intColCount];
        intGroupType = 1;
        parseGroupTemplate();
    }

    public Group(String strGroupFieldName)
        throws FormulaException
    {
        sheet = null;
        rs = null;
        this.strGroupFieldName = "";
        strPGroupValue = "(--)";
        strGroupValue = "";
        intPRow = 0;
        intCRow = 0;
        intDetailX = 0;
        intTemplateRow = 0;
        intColCount = 0;
        vtTemplateCell = null;
        vtCellType = null;
        vtTemplateMerge = null;
        vtTemplateFormula = null;
        vtTemplateFormula1 = null;
        vtFieldList = null;
        vtFieldValueList = null;
        intGroupType = 0;
        footer = null;
        this.strGroupFieldName = strGroupFieldName;
    }

    public int getTemplateRow()
    {
        return intTemplateRow;
    }

    public String getGroupID()
    {
        return strGroupFieldName;
    }

    public void increaseColCount()
    {
        intColCount++;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount];
        vtCellType = new int[intColCount];
        vtFieldValueList = new String[intColCount];
    }

    public void setSheet(WritableSheet sheet)
    {
        this.sheet = sheet;
    }

    private void parseGroupTemplate()
        throws FormulaException, Exception
    {
        WritableCell cell = sheet.getWritableCell(0, intTemplateRow);
        if(cell.getCellFeatures() != null)
        {
            strGroupFieldName = replace(replace(cell.getCellFeatures().getComment(), "\n", ""), " ", "");
            if(!FieldExists(rs, strGroupFieldName))
            {
                throw new Exception("Field " + strGroupFieldName + " that is group id at row " + (intTemplateRow + 1) + " does not exists");
            }
            cell.getCellFeatures().removeComment();
        } else
        {
            throw new Exception("Group part whose template at row " + (intTemplateRow + 1) + " has not group id");
        }
        for(int inti = 0; inti < intColCount; inti++)
        {
            Cell cell1 = sheet.getCell(inti + intDetailX, intTemplateRow);
            vtTemplateCell[inti] = cell1;
            vtCellType[inti] = Util.getCellType(cell1);
            if(cell1.toString().toLowerCase().indexOf("formula") >= 0)
            {
                FormulaCell fc = (FormulaCell)cell1;
                String strFormula = fc.getFormula();
                ReportTDK.Formula formula = new ReportTDK.Formula(cell1, strFormula);
                if(strFormula.indexOf("${") >= 0)
                {
                    vtTemplateFormula1.add(formula);
                } else
                {
                    vtTemplateFormula.add(formula);
                }
                vtFieldValueList[inti] = "";
                vtCellType[inti] = 3;
                continue;
            }
            String cv = cell1.getContents();
            vtFieldValueList[inti] = cv;
            int intn = -1;
            do
            {
                int intm = cv.indexOf("${", intn + 1);
                if(intm < 0)
                {
                    break;
                }
                intn = cv.indexOf("}", intm);
                if(intn < 0)
                {
                    break;
                }
                String strFieldName = cv.substring(intm + 2, intn).trim();
                if(!FieldExists(rs, strFieldName))
                {
                    throw new Exception("Error: Field " + strFieldName + " that be needed to fill at cell(" + (inti + intDetailX + 1) + ", " + (intTemplateRow + 1) + " ) does not exists");
                }
                vtFieldList[inti].add(strFieldName);
            } while(true);
            if(vtFieldList[inti].size() > 0 && vtCellType[inti] == 0)
            {
                throw new Exception("Cell( " + (inti + 1) + ", " + (intTemplateRow + 1) + ") must to be formated as text or numeric.");
            }
        }

        Range rl1[] = sheet.getMergedCells();
        for(int inti1 = 0; inti1 < rl1.length; inti1++)
        {
            Range r1 = rl1[inti1];
            Cell cell1 = r1.getTopLeft();
            Cell cell2 = r1.getBottomRight();
            if(cell1.getRow() == cell2.getRow() && cell1.getRow() == intTemplateRow)
            {
                vtTemplateMerge.add(new Merge(cell1.getColumn(), intTemplateRow, cell2.getColumn(), intTemplateRow));
                sheet.unmergeCells(r1);
            }
        }

    }

    public void reParseDetailTemplate()
        throws FormulaException, SQLException, Exception
    {
        for(int inti = 0; inti < intColCount; inti++)
        {
            Cell cell1 = sheet.getCell(inti + intDetailX, intTemplateRow);
            vtTemplateCell[inti] = cell1;
            vtCellType[inti] = Util.getCellType(cell1);
            if(cell1.toString().toLowerCase().indexOf("formula") == -1)
            {
                String cv = cell1.getContents();
                vtFieldValueList[inti] = cv;
                vtFieldList[inti].clear();
                int intn = -1;
                do
                {
                    int intm = cv.indexOf("${", intn + 1);
                    if(intm < 0)
                    {
                        break;
                    }
                    intn = cv.indexOf("}", intm);
                    if(intn < 0)
                    {
                        break;
                    }
                    String strFieldName = cv.substring(intm + 2, intn);
                    if(!FieldExists(rs, strFieldName))
                    {
                        throw new Exception("Error: Field " + strFieldName + " that be needed to fill at cell(" + (inti + intDetailX + 1) + ", " + (intTemplateRow + 1) + " ) does not exists");
                    }
                    vtFieldList[inti].add(strFieldName);
                } while(true);
                if(vtFieldList[inti].size() > 0 && vtCellType[inti] == 0)
                {
                    throw new Exception("Cell( " + (inti + 1) + ", " + (intTemplateRow + 1) + ") must to be formated as text or numeric.");
                }
            } else
            {
                vtFieldValueList[inti] = "";
                vtCellType[inti] = 3;
            }
        }

    }

    public void clearPreviousGroupID()
    {
        strPGroupValue = "";
    }

    public boolean isGroupChanged()
        throws SQLException
    {
        strGroupValue = rs.getObject(strGroupFieldName).toString();
        return !strPGroupValue.equals(strGroupValue);
    }

    public boolean fillFooter(int intRow)
        throws SQLException, WriteException
    {
        if((intGroupType & 2) != 2)
        {
            return false;
        }
        if(strPGroupValue.equals(""))
        {
            return false;
        }
        if(intPRow == 0)
        {
            return false;
        } else
        {
            footer.fillData(intRow);
            footer.setFirstPos(intRow + 1);
            return true;
        }
    }

    public boolean fillGroup(int intRow)
        throws SQLException, WriteException
    {
        strPGroupValue = strGroupValue;
        intPRow = intCRow;
        intCRow = intRow;
        if((intGroupType & 1) != 1)
        {
            return false;
        }
        sheet.insertRow(intRow);
        addMergeToSheet(sheet, vtTemplateMerge, intRow);
        for(int intx = 0; intx < intColCount; intx++)
        {
            String strValue = getCellValueToFill(rs, vtFieldList[intx], vtFieldValueList[intx]);
            Util.addCellToSheet(sheet, vtTemplateCell[intx], vtCellType[intx], strValue, intx + intDetailX, intRow);
        }

        fillFormula1(intRow);
        return true;
    }

    public void fillFormula(int intRow)
        throws WriteException, SQLException
    {
        if(intCRow == intPRow)
        {
            return;
        }
        for(int inti = 0; inti < vtTemplateFormula.size(); inti++)
        {
            ReportTDK.Formula f = (ReportTDK.Formula)vtTemplateFormula.get(inti);
            sheet.addCell(new jxl.write.Formula(f.getC(), intPRow, f.getFormula(intPRow + 2, intRow), f.getCellFormat()));
        }

    }

    public void fillFormula1(int intRow)
        throws WriteException, SQLException
    {
        for(int inti = 0; inti < vtTemplateFormula1.size(); inti++)
        {
            ReportTDK.Formula f = (ReportTDK.Formula)vtTemplateFormula1.get(inti);
            sheet.addCell(new jxl.write.Formula(f.getC(), intRow, f.getFormula(rs), f.getCellFormat()));
        }

    }

    public void addColumnGroup(int intCol, int intTemplateCol, String strValue)
    {
        for(int inti = 0; inti < vtTemplateFormula.size(); inti++)
        {
            ReportTDK.Formula formula = (ReportTDK.Formula)vtTemplateFormula.get(inti);
            formula.strFormula = replace(formula.strFormula, "E#", "F#");
            formula.strFormula1 = replace(formula.strFormula1, "E#", "F#");
        }

    }

    public void prepairFormulas(int intCol)
    {
        prepairFormulas(intCol, vtTemplateFormula);
        prepairFormulas(intCol, vtTemplateFormula1);
    }

    private void prepairFormulas(int intCol, Vector vtFormula)
    {
        for(int inti = 0; inti < vtFormula.size(); inti++)
        {
            ReportTDK.Formula formula = (ReportTDK.Formula)vtFormula.get(inti);
            formula.prepairFormula(sheet, intCol);
        }

    }

    public void prepairMerges(int intCol)
    {
        for(int inti = 0; inti < vtTemplateMerge.size(); inti++)
        {
            Merge merge = (Merge)vtTemplateMerge.get(inti);
            merge.prepairMerge(intCol);
        }

    }

    public void addFormula(int intCol, String strFormula)
    {
        Cell cell = sheet.getCell(intCol, intTemplateRow);
        ReportTDK.Formula formula = new ReportTDK.Formula(cell, strFormula);
        if(strFormula.indexOf("${") >= 0)
        {
            vtTemplateFormula1.add(formula);
        } else
        {
            vtTemplateFormula.add(formula);
        }
        vtFieldValueList[intCol] = "";
        vtCellType[intCol] = 3;
    }

    public int getGroupType()
    {
        return intGroupType;
    }

    public void setGroupType(int intGroupType)
    {
        this.intGroupType = intGroupType;
    }

    public Total getGroupFooter()
    {
        return footer;
    }

    public void setGroupFooter(Total footer)
    {
        this.footer = footer;
        intGroupType = intGroupType | 2;
    }
    public void setResultSet(ResultSet rs)
    {
        this.rs= rs;
    }
    public void setResultSet(WritableSheet sheet)
    {
        this.sheet= sheet;
    }

    public void unInit()
    {
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti].clear();
        }

        vtFieldList = null;
        vtTemplateCell = null;
        vtTemplateMerge.clear();
        vtTemplateMerge = null;
        vtTemplateFormula.clear();
        vtTemplateFormula = null;
        vtTemplateFormula1.clear();
        vtTemplateFormula1 = null;
        vtFieldValueList = null;
        if(footer != null)
        {
            footer.unInit();
        }
    }
}
