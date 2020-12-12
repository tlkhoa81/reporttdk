package ReportTDK;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Vector;
import jxl.*;
import jxl.biff.formula.FormulaException;
import jxl.write.*;

/**
 * <p>Title: ReportTDK</p>
 *
 * <p>Description: Library to create excel report from template</p>
 *
 * @version 2.0
 */

class Total
{

    private ResultSet rs;
    private WritableSheet sheet;
    private int intTemplateRow;
    private int intTemplateRow1;
    private int intColCount;
    private Cell vtTemplateCell[];
    public int vtCellType[];
    private Vector vtTemplateMerge;
    private Vector vtTemplateFormula;
    private Vector vtFieldList[];
    private String vtFieldValueList[];
    private int intDetailX;
    private int intTemplateHeight;
    private int intGroupFooterCount;
    private String strGroupFieldName;

    public Total(WritableSheet sheet, ResultSet rs, int intTemplateRow, int intColCount, int intDetailX)
        throws FormulaException, SQLException, Exception
    {
        this.rs = null;
        this.sheet = null;
        this.intTemplateRow = 0;
        intTemplateRow1 = 0;
        this.intColCount = 1;
        vtTemplateCell = null;
        vtCellType = null;
        vtTemplateMerge = null;
        vtTemplateFormula = null;
        vtFieldList = null;
        vtFieldValueList = null;
        this.intDetailX = 0;
        intTemplateHeight = 0;
        intGroupFooterCount = 0;
        strGroupFieldName = "";
        intTemplateRow--;
        this.sheet = sheet;
        this.rs = rs;
        intColCount++;
        this.intTemplateRow = intTemplateRow;
        intTemplateRow1 = intTemplateRow;
        this.intDetailX = intDetailX;
        this.intColCount = intColCount;
        intTemplateHeight = intTemplateHeight;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount + 1];
        vtCellType = new int[intColCount + 1];
        vtTemplateMerge = new Vector();
        vtTemplateFormula = new Vector();
        vtFieldValueList = new String[intColCount];
        parseTemplate();
    }

    public int getTemplateRow()
    {
        return intTemplateRow;
    }

    public int getTemplateRow1()
    {
        return intTemplateRow1;
    }

    public void setTemplateRow1(int intTemplateRow1)
    {
        this.intTemplateRow1 = intTemplateRow1;
    }

    public void increaseColCount()
    {
        intColCount++;
        vtFieldList = new Vector[intColCount];
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti] = new Vector();
        }

        vtTemplateCell = new Cell[intColCount + 1];
        vtCellType = new int[intColCount + 1];
        vtFieldValueList = new String[intColCount];
    }

    public void setSheet(WritableSheet sheet)
    {
        this.sheet = sheet;
    }

    private void parseTemplate()
        throws FormulaException, SQLException, Exception
    {
        WritableCell cell = sheet.getWritableCell(0, intTemplateRow);
        if(cell.getCellFeatures() != null)
        {
            strGroupFieldName = cell.getCellFeatures().getComment();
            cell.getCellFeatures().removeComment();
        }
        for(int inti = 0; inti < intColCount; inti++)
        {
            Cell cell1 = sheet.getCell(inti + intDetailX, intTemplateRow);
            vtTemplateCell[inti] = cell1;
            vtCellType[inti] = Util.getCellType(cell1);
            if(cell1.toString().toLowerCase().indexOf("formula") >= 0)
            {
                FormulaCell fc = (FormulaCell)cell1;
                ReportTDK.Formula formula = new ReportTDK.Formula(cell1, fc.getFormula());
                vtTemplateFormula.add(formula);
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
                String strFieldName = cv.substring(intm + 2, intn);
                if(!Group.FieldExists(rs, strFieldName))
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

    public void reParseTemplate()
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
                    if(!Group.FieldExists(rs, strFieldName))
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

    public void fillData(int intRow)
        throws WriteException, SQLException
    {
        sheet.insertRow(intRow);
        Group.addMergeToSheet(sheet, vtTemplateMerge, intRow);
        for(int intx = 0; intx < intColCount; intx++)
        {
            String strValue = Group.getCellValueToFill(rs, vtFieldList[intx], vtFieldValueList[intx]);
            Util.addCellToSheet(sheet, vtTemplateCell[intx], vtCellType[intx], strValue, intx + intDetailX, intRow);
        }

        fillFormula(intRow);
    }

    public void fillFormula(int intRow)
        throws WriteException, SQLException
    {
        for(int inti = 0; inti < vtTemplateFormula.size(); inti++)
        {
            ReportTDK.Formula f = (ReportTDK.Formula)vtTemplateFormula.get(inti);
            sheet.addCell(new jxl.write.Formula(f.getC(), intRow, f.getFormula(intRow + 1, intTemplateRow1 - intGroupFooterCount, intRow), f.getCellFormat()));
        }

    }

    public void prepairFormulas(int intCol)
    {
        prepairFormulas(intCol, vtTemplateFormula);
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
        vtTemplateFormula.add(formula);
        vtFieldValueList[intCol] = "";
        vtCellType[intCol] = 3;
    }

    public String getGroupID()
    {
        return strGroupFieldName;
    }

    public void setGroupID(String strGroupFieldName)
    {
        this.strGroupFieldName = strGroupFieldName;
    }

    public void setFirstPos(int intFirstPos)
    {
        intTemplateRow1 = intFirstPos;
    }

    public void setCellValue(int intCol, String strValue)
    {
        vtFieldValueList[intCol - 1] = strValue;
    }

    public void setGroupFooterCount(int intGroupFooterCount)
    {
        this.intGroupFooterCount = intGroupFooterCount;
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
        if (vtFieldList!=null)
        for(int inti = 0; inti < intColCount; inti++)
        {
            vtFieldList[inti].clear();
        }

        vtFieldList = null;
        vtTemplateCell = null;
        if (vtTemplateMerge!=null) vtTemplateMerge.clear();
        vtTemplateMerge = null;
        if (vtTemplateFormula!=null) vtTemplateFormula.clear();
        vtTemplateFormula = null;
        vtFieldValueList = null;
    }
}
